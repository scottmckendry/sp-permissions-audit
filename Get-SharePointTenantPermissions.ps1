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
    [string] $CertificatePath
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
                Write-Error "Failed to connect to $SiteUrl after $connectionAttempts attempts."
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

    return Get-MsalToken @connectionParameters
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
    $groupMemberShipResponse = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/users/$UserEmail/memberOf" -Method GET -Headers @{
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

    $groupMembership = @()
    foreach ($group in $groupMemberShipResponse.value) {
        $groupMembership += [PSCustomObject]@{
            GroupName = $group.displayName
            GroupId   = $group.id
        }
    }

    return $groupMembership
}

function Test-UserIsSiteCollectionAdmin {
    <#
    .SYNOPSIS
    Checks if a given user is a site collection admin for a given site collection.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $UserEmail
    )

    $siteAdmins = Get-PnPSiteCollectionAdmin
    foreach ($siteAdmin in $siteAdmins) {
        $siteAdminLogin = $siteAdmin.LoginName.Split('|')[2]

        if ($UserEmail -eq $siteAdminLogin) {
            return $true
        }

        # Check if user is a member of a group that is a site collection admin
        if ($userGroupMembership.GroupId -contains $siteAdminLogin) {
            return $true
        }
    }

    return $false
}

function Test-UserInSharePointGroup {
    <#
    .SYNOPSIS
    Returns an array of SharePoint groups that a given user is a member of for a given site collection.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $UserEmail,
        [Parameter(Mandatory = $true)]
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
                $groupMembership += [PSCustomObject]@{
                    GroupName       = $siteGroup.Title
                    PermissionLevel = $groupPermissionLevel.Name
                }
            }
            elseif ($userGroupMembership.GroupId -contains $groupMemberLogin) {
                $groupPermissionLevel = Get-PnPGroupPermissions -Identity $siteGroup
                $groupMembership += [PSCustomObject]@{
                    GroupName       = $siteGroup.Title
                    PermissionLevel = $groupPermissionLevel.Name
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
        [string] $UserSharePointId
    )

    $ctx = Get-PnPContext
    $web = $ctx.Web
    $ctx.Load($web)
    $ctx.ExecuteQuery()

    $lists = $web.Lists
    $ctx.Load($lists)
    $ctx.ExecuteQuery()

    # Exlude built-in lists
    $excludedLists = @("App Packages", "appdata", "appfiles", "Apps in Testing", "Cache Profiles", "Composed Looks", "Content and Structure Reports", "Content type publishing error log", "Converted Forms", "Device Channels", "Form Templates", "fpdatasources", "Get started with Apps for Office and SharePoint", "List Template Gallery", "Long Running Operation Status", "Maintenance Log Library", "Style Library", , "Master Docs", "Master Page Gallery", "MicroFeed", "NintexFormXml", "Quick Deploy Items", "Relationships List", "Reusable Content", "Search Config List", "Solution Gallery", "Site Collection Images", "Suggested Content Browser Locations", "TaxonomyHiddenList", "User Information List", "Web Part Gallery", "wfpub", "wfsvc", "Workflow History", "Workflow Tasks", "Preservation Hold Library")
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
            Write-Host "$(Get-Date) INFO: `tGetting permissions for $($list.Title) and user $UserSharePointId..."
            $listPermissions = Get-PnPListPermissions -PrincipalId $UserSharePointId -Identity $list.Title

            $siteListPermissions += [PSCustomObject]@{
                ListName        = $list.Title
                PermissionLevel = $listPermissions.Name
            }
        }
    }

    return $siteListPermissions
}

Write-Host "$(Get-Date) INFO: Connecting to tenant admin site..."
Connect-TenantSite -SiteUrl "https://$TenantName-admin.sharepoint.com"

Write-Host "$(Get-Date) INFO: Getting all site collections..."
$siteCollections = Get-PnPTenantSite
Write-Host "$(Get-Date) INFO: `tFound $($siteCollections.Count) site collections."
Disconnect-PnPOnline

Write-Host "$(Get-Date) INFO: Getting group membership for $UserEmail..."
$userGroupMembership = Get-UserGroupMembership -UserEmail $UserEmail
Write-Host "$(Get-Date) INFO: `tFound $($userGroupMembership.Count) groups."

# Create CSV if doesn't exist, remove if does and recreate
if (Test-Path $CSVPath) {
    Remove-Item $CSVPath
}
New-Item -Path $CSVPath -ItemType File | Out-Null

#Add CSV Header row
Add-Content -Path $CSVPath -Value "User,Site URL,List/Library,Group,Permission Level"

$siteCounter = 1
$userId = 0
foreach ($siteCollection in $siteCollections) {
    Write-Host "$(Get-Date) INFO: Connecting to $($siteCollection.Url)`t ($siteCounter of $($siteCollections.Count))..."
    $siteCounter++
    Connect-TenantSite -SiteUrl $siteCollection.Url

    $user = Get-PnPUser | Where-Object Email -EQ $UserEmail
    if ($user) {
        $userId = $user.Id
    }

    if (Test-UserIsSiteCollectionAdmin -UserEmail $UserEmail) {
        Write-Host "$(Get-Date) INFO: `t$UserEmail is a site collection admin for $($siteCollection.Url)."
        Add-Content -Path $CSVPath -Value "$UserEmail, $($siteCollection.Url), , Site Admins, Full Control"
        Disconnect-PnPOnline
        continue
    }

    $sharepointGroupMembership = Test-UserInSharePointGroup -UserEmail $UserEmail -GraphGroups $userGroupMembership
    if ($sharepointGroupMembership) {
        foreach ($group in $sharepointGroupMembership) {
            Write-Host "$(Get-Date) INFO: `t$UserEmail is a member of $($group.GroupName) with $($group.PermissionLevel) permissions."
            Add-Content -Path $CSVPath -Value "$UserEmail, $($siteCollection.Url), , $($group.GroupName), $($group.PermissionLevel)"
        }
    }

    # Write-Host "$(Get-Date) INFO: `tChecking unique permissions for $($siteCollection.Url)..." -ForegroundColor Cyan
    if ($userId -ne 0) {
        $listPermissions = Get-UniqueListPermissions -UserSharePointId $userId
        if ($listPermissions.Count -gt 0) {
            foreach ($listPermission in $listPermissions) {
                # Write-Host "$(Get-Date) INFO: `t$UserEmail has $($listPermission.PermissionLevel) permissions on $($listPermission.ListName)."
                Add-Content -Path $CSVPath -Value "$UserEmail, $($siteCollection.Url), $($listPermission.ListName), , $($listPermission.PermissionLevel)"
            }
        }
    }

    Disconnect-PnPOnline
    $userId = 0
}

