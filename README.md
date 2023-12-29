# SharePoint Online Permissions Audit Script

It is well known that SharePoint permissions are notoriously difficult to manage. This script is designed to help you audit permissions across your SharePoint Online sites.

## ✨ Features

-   Audit permissions for all sites in a SharePoint Online tenant - all the way down to list and library level.
-   Capture permissions granted to Security (Entra ID) and Microsoft 365 groups.
-   Uses a modern authentication flow that does not require a user to be logged in or have access to all sites in the tenant.

## 📝 Output

The script will output a CSV file with the following columns:

| Column Name       | Description                                                                                         |
| ----------------- | --------------------------------------------------------------------------------------------------- |
| UserPrincipalName | The user's UPN/email address                                                                        |
| SiteUrl           | The URL of the site                                                                                 |
| SiteAdmin         | Is the user a site admin?                                                                           |
| GroupName         | If the user is not a site admin, what SharePoint group are they in?                                 |
| PermissionLevel   | The permission level granted to the SharePoint group, e.g full control, read, edit etc.             |
| ListName          | The title of a list or library where the user has unique permissions. (also captures sharing links) |
| ListPermission    | The permission level granted to the user on the list or library.                                    |

## 🚀 Getting Started

### Prerequisites

-   Global Adminstrator Role
-   PowerShell 7 or later with the latest versions of [PnP.PowerShell](https://pnp.github.io/powershell/) and [MSAL.PS](https://github.com/AzureAD/MSAL.PS/) modules installed.
-   A self-signed certificate for use with the app registration. See [this article](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread) for more information.

```powershell
Install-Module -Name PnP.PowerShell -Scope CurrentUser
Install-Module -Name MSAL.PS -Scope CurrentUser
```

### Create an Entra ID App Registration

Follow the steps in [this article](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread) to create an app registration in Azure AD. Make sure you grant the app the following permissions.

**Graph API**

-   Sites.Read.All
-   Directory.Read.All

**SharePoint API**

-   Sites.Read.All

## Usage

The intention is for this script to be called by a parent script that will pass in the required parameters. This allows you to run the script against multiple users and potentially multiple tenants.
Below is an example of how you might call the script.

```powershell
# audit.ps1 - Create in the same directory as Get-SharePointOnlinePermissions.ps1

$tenantName = "contoso" # The name of your tenant, e.g. contoso.sharepoint.com
$csvPath = "C:\temp\permissions.csv" # The path to the output CSV file
$clientID = "00000000-0000-0000-0000-000000000000" # The client ID of the app registration
$certificatePath = "C:\temp\certificate.pfx" # The path to the certificate filer
$append = $true # Should the script append to the CSV file or overwrite it?

$users = @(
    "john@contoso.com",
    "jane@contoso.com"
)

foreach ($user in $users) {
    .\Get-SharePointOnlinePermissions.ps1 -TenantName $tenantName -CsvPath $csvPath -ClientID $clientID -CertificatePath $certificatePath -Append $append -UserEmail $user
}

```

## 🤝 Contributing

Contributions, issues and feature requests are welcome!

TODO:

-   [ ] Replace [MSAL.PS](https://github.com/AzureAD/MSAL.PS) cmdlets with a non-deprecated alternative
