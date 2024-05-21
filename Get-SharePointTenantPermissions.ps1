# Get-SharePointTenantPermissions.ps1
# Description: This script will get all the permissions for a given user or users in a SharePoint Online tenant and export them to a CSV file.

#requires -Modules PnP.PowerShell, MSAL.PS
param (
    [Parameter(Mandatory = $true)]
    [string] $TenantName,
    [Parameter(Mandatory = $true)]
    [string] $UserEmail,
    [Parameter(Mandatory = $true)]
    [string] $CSVPath,
    [Parameter(Mandatory = $true)]
    [string] $ClientId,
    [Parameter(Mandatory = $true)]
    [string] $CertificatePath,
    [Parameter(Mandatory = $false)]
    [switch] $Append = $false
)

function Connect-TenantSite {
    <#
    .SYNOPSIS
    Connects to a SharePoint Online site using certificate-based authentication via PnP PowerShell.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl
    )

    $connectionAttempts = 3
    for ($i = 0; $i -lt $connectionAttempts; $i++) {
        try {
            Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -CertificatePath $CertificatePath -Tenant "$TenantName.onmicrosoft.com"
            break
        }
        catch {
            if ($i -eq $connectionAttempts - 1) {
                Write-Error $_.Exception.Message
                throw $_
            }
            continue
        }
    }
}

function Get-GraphToken {
    <#
    .SYNOPSIS
    Gets a bearer token for the Microsoft Graph API using certificate-based authentication.
    #>
    $connectionParameters = @{
        'TenantId'          = "$TenantName.onmicrosoft.com"
        'ClientId'          = $ClientId
        'ClientCertificate' = $CertificatePath
    }

    try {
        return Get-MsalToken @connectionParameters
    }
    catch {
        Write-Error $_.Exception.Message
        throw $_
    }
}

function Get-UserGroupMembership {
    <#
    .SYNOPSIS
    Gets the group membership for a given user. Returns an array of objects containing the group name and group id.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $UserEmail
    )

    $accessToken = Get-GraphToken
    $encodedUserEmail = [System.Web.HttpUtility]::UrlEncode($UserEmail)

    try {
        $groupMemberShipResponse = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/users/$encodedUserEmail/memberOf" -Method GET -Headers @{
            Authorization = "Bearer $($accessToken.AccessToken)"
        } | ConvertFrom-Json

        # If @odata.nextLink exists, get next page of results
        while ($groupMemberShipResponse.'@odata.nextLink') {
            $appendGroupMembershipResponse = Invoke-WebRequest -Uri $groupMemberShipResponse.'@odata.nextLink' -Method GET -Headers @{
                Authorization = "Bearer $($accessToken.AccessToken)"
            }
            $graphGroupMembership.value += $appendGroupMembershipResponse.value
            $graphGroupMembership.'@odata.nextLink' = $appendGroupMembershipResponse.'@odata.nextLink'
        }
    }
    catch {
        Write-Error $_.Exception.Message
        throw $_
    }

    $groupMembership = @()
    foreach ($group in $groupMemberShipResponse.value) {
        $groupMembership += [PSCustomObject]@{
            GroupName = $group.displayName
            GroupId   = $group.id
        }
    }

    return $groupMembership
}

function New-CsvFile {
    <#
    .SYNOPSIS
    Creates a new CSV file.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $csv = [PSCustomObject]@{
        UserPrincipalName = $null
        SiteUrl           = $null
        SiteAdmin         = $null
        GroupName         = $null
        PermissionLevel   = $null
        ListName          = $null
        ListPermission    = $null
    }

    if (Test-Path $Path) {
        Remove-Item $Path
    }

    $csv | Export-Csv -Path $Path -NoTypeInformation

    # Remove the first (empty) line of the CSV file
    $csvFile = Get-Content $Path
    $csvFile = $csvFile[0..($csvFile.Length - 2)]
    Set-Content -Path $Path -Value $csvFile
}

function Test-UserIsSiteCollectionAdmin {
    <#
    .SYNOPSIS
    Checks if a given user is a site collection admin for a given site collection.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $UserEmail,
        [Parameter(Mandatory = $false)]
        [array] $GraphGroups
    )

    $siteAdmins = Get-PnPSiteCollectionAdmin
    foreach ($siteAdmin in $siteAdmins) {
        $siteAdminLogin = $siteAdmin.LoginName.Split('|')[2]

        if ($UserEmail -eq $siteAdminLogin) {
            return $true
        }

        # Check if user is a member of a group that is a site collection admin
        if ($null -ne $GraphGroups) {
            if ($userGroupMembership.GroupId -contains $siteAdminLogin) {
                return $true
            }
        }
    }

    return $false
}

function Get-UserSharePointGroups {
    <#
    .SYNOPSIS
    Returns an array of SharePoint groups that a given user is a member of for a given site collection.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $UserEmail,
        [Parameter(Mandatory = $false)]
        [array] $GraphGroups
    )

    $siteGroups = Get-PnPGroup

    $groupMembership = @()
    foreach ($siteGroup in $siteGroups) {
        $groupMembers = Get-PnPGroupMember -Identity $siteGroup.Title

        foreach ($groupMember in $groupMembers) {
            $groupMemberLogin = $groupMember.LoginName.Split('|')[2]
            if ($UserEmail -eq $groupMemberLogin) {
                $groupPermissionLevel = Get-PnPGroupPermissions -Identity $siteGroup
                $permissionLevelString = ""
                foreach ($permissionLevel in $groupPermissionLevel) {
                    $permissionLevelString += $permissionLevel.Name + " | "
                }

                if ($permissionLevelString -eq "") {
                    $permissionLevelString = "No Permissions"
                }
                else {
                    # remove trailing " | "
                    $permissionLevelString = $permissionLevelString.Substring(0, $permissionLevelString.Length - 3)
                }

                $groupMembership += [PSCustomObject]@{
                    GroupName       = $siteGroup.Title
                    PermissionLevel = $permissionLevelString
                }

            }
            elseif ($null -ne $GraphGroups) {
                if ($userGroupMembership.GroupId -contains $groupMemberLogin) {
                    $groupPermissionLevel = Get-PnPGroupPermissions -Identity $siteGroup
                    $permissionLevelString = ""
                    foreach ($permissionLevel in $groupPermissionLevel) {
                        $permissionLevelString += $permissionLevel.Name + " | "
                    }

                    if ($permissionLevelString -eq "") {
                        $permissionLevelString = "No Permissions"
                    }
                    else {
                        # remove trailing " | "
                        $permissionLevelString = $permissionLevelString.Substring(0, $permissionLevelString.Length - 3)
                    }

                    $groupMembership += [PSCustomObject]@{
                        GroupName       = $siteGroup.Title
                        PermissionLevel = $permissionLevelString
                    }
                }
            }
        }
    }

    return $groupMembership
}

function Get-UniqueListPermissions {
    <#
    .SYNOPSIS
    Gets the unique permissions at the list level for a given user for a given site collection.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $UserEmail,
        [Parameter(Mandatory = $false)]
        [array] $GraphGroups
    )

    $ctx = Get-PnPContext
    $web = $ctx.Web
    $ctx.Load($web)
    $ctx.ExecuteQuery()

    $lists = $web.Lists
    $ctx.Load($lists)
    $ctx.ExecuteQuery()

    # Exlude built-in lists
    $excludedLists = @("App Packages", "appdata", "appfiles", "Apps in Testing", "Cache Profiles", "Composed Looks", "Content and Structure Reports", "Content type publishing error log", "Converted Forms", "Device Channels", "Form Templates", "fpdatasources", "Get started with Apps for Office and SharePoint", "List Template Gallery", "Long Running Operation Status", "Maintenance Log Library", "Style Library", , "Master Docs", "Master Page Gallery", "MicroFeed", "NintexFormXml", "Quick Deploy Items", "Relationships List", "Reusable Content", "Search Config List", "Solution Gallery", "Site Collection Images", "Suggested Content Browser Locations", "TaxonomyHiddenList", "User Information List", "Web Part Gallery", "wfpub", "wfsvc", "Workflow History", "Workflow Tasks", "Preservation Hold Library", "SharePointHomeCacheList")
    $siteListPermissions = @()
    foreach ($list in $lists) {
        $ctx.Load($list)
        $ctx.ExecuteQuery()

        if ($excludedLists -contains $list.Title) {
            continue
        }

        $list.Retrieve("HasUniqueRoleAssignments")
        $ctx.ExecuteQuery()

        if ($list.HasUniqueRoleAssignments) {
            $listPermissions = $list.RoleAssignments
            $ctx.Load($listPermissions)
            $ctx.ExecuteQuery()

            foreach ($roleassignment in $listPermissions) {
                $ctx.Load($roleassignment.Member)
                $ctx.Load($roleassignment.RoleDefinitionBindings)
                $ctx.ExecuteQuery()

                if ($UserEmail -eq ($roleassignment.Member.LoginName.Split('|')[2])) {
                    $listPermission = [PSCustomObject]@{
                        Name            = $list.Title
                        PermissionLevel = $roleassignment.RoleDefinitionBindings.Name
                    }

                    $siteListPermissions += $listPermission
                }
                elseif ($null -ne $GraphGroups) {
                    if ( $GraphGroups.GroupId -contains ($roleassignment.Member.LoginName.Split('|')[2])) {
                        $listPermission = [PSCustomObject]@{
                            Name            = $list.Title
                            PermissionLevel = $roleassignment.RoleDefinitionBindings.Name
                        }

                        $siteListPermissions += $listPermission
                    }
                }
            }
        }
    }
    return $siteListPermissions
}

Set-Location $PSScriptRoot

Write-Host "$(Get-Date) INFO: Connecting to tenant admin site..."
Connect-TenantSite -SiteUrl "https://$TenantName-admin.sharepoint.com" -ErrorAction Stop

Write-Host "$(Get-Date) INFO: Getting all site collections..."
$siteCollections = Get-PnPTenantSite -ErrorAction Stop
Write-Host "$(Get-Date) INFO: `tFound $($siteCollections.Count) site collections."
Disconnect-PnPOnline

Write-Host "$(Get-Date) INFO: Getting group membership for $UserEmail..."
$userGroupMembership = Get-UserGroupMembership -UserEmail $UserEmail -ErrorAction Stop
Write-Host "$(Get-Date) INFO: `tFound $($userGroupMembership.Count) groups."

if (!$Append) {
    New-CsvFile -Path $CSVPath
}

$siteCounter = 1
foreach ($siteCollection in $siteCollections) {
    Write-Host "$(Get-Date) INFO: Connecting to $($siteCollection.Url)`t ($siteCounter of $($siteCollections.Count))..."
    $siteCounter++
    Connect-TenantSite -SiteUrl $siteCollection.Url

    if (Test-UserIsSiteCollectionAdmin -UserEmail $UserEmail) {
        Write-Host "$(Get-Date) INFO: `t$UserEmail is a site collection admin for $($siteCollection.Url)."
        $csvLineObject = [PSCustomObject]@{
            UserPrincipalName = $UserEmail
            SiteUrl           = $siteCollection.Url
            SiteAdmin         = $true
            GroupName         = $null
            PermissionLevel   = $null
            ListName          = $null
            ListPermission    = $null
        }
        $csvLineObject | Export-Csv -Path $CSVPath -Append -NoTypeInformation
        continue
    }

    # Check if user is a member of any SharePoint groups
    $sharepointGroupMembership = Get-UserSharePointGroups -UserEmail $UserEmail -GraphGroups $userGroupMembership
    if ($sharepointGroupMembership) {
        foreach ($group in $sharepointGroupMembership) {
            Write-Host "$(Get-Date) INFO: `t$UserEmail is a member of $($group.GroupName) with $($group.PermissionLevel) permissions."
            $csvLineObject = [PSCustomObject]@{
                UserPrincipalName = $UserEmail
                SiteUrl           = $siteCollection.Url
                SiteAdmin         = $false
                GroupName         = $group.GroupName
                PermissionLevel   = $group.PermissionLevel
                ListName          = $null
                ListPermission    = $null
            }
            $csvLineObject | Export-Csv -Path $CSVPath -Append -NoTypeInformation
        }
    }

    # Check if user has unique permissions at the list level
    $listPermissions = Get-UniqueListPermissions -UserEmail $UserEmail -GraphGroups $userGroupMembership
    if ($listPermissions.Count -gt 0) {
        foreach ($listPermission in $listPermissions) {
            Write-Host "$(Get-Date) INFO: `t$UserEmail has $($listPermission.PermissionLevel) permissions on $($listPermission.Name)."
            $csvLineObject = [PSCustomObject]@{
                UserPrincipalName = $UserEmail
                SiteUrl           = $siteCollection.Url
                SiteAdmin         = $false
                GroupName         = $null
                PermissionLevel   = $null
                ListName          = $listPermission.Name
                ListPermission    = $listPermission.PermissionLevel
            }
            $csvLineObject | Export-Csv -Path $CSVPath -Append -NoTypeInformation
        }
    }

    # Reset user ID - Prevents false positives and extra work being done on sites the user has never visited
    Disconnect-PnPOnline
}
