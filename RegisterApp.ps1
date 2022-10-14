# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT license.

param(
  [Parameter(Mandatory=$false,
  HelpMessage="The friendly name of the app registration")]
  [String]
  $AppName = "My Sample Teams App",

  [Parameter(Mandatory=$true,
  HelpMessage="The fully-qualified domain of your web API. If you are running locally, this should be your ngrok forwarding URL.")]
  [String]
  $AppDomain,

  [Parameter(Mandatory=$false)]
  [Switch]
  $UseDeviceAuthentication = $false,

  [Parameter(Mandatory=$false)]
  [Switch]
  $StayConnected = $false
)

$GraphAppId = "00000003-0000-0000-c000-000000000000"

# Validate and normalize AppDomain
# Check for and remove any protocol specifier (http://, https://)
# and any path after the host
if ($AppDomain.Contains("://"))
{
    $SubIndex = $AppDomain.IndexOf("://") + 3
    $AppDomain = $AppDomain.Substring($SubIndex)
    $AppDomain = $AppDomain.Split('/')[0]

    # Validate the result
    if ($null -eq (("https://" + $AppDomain) -as [System.Uri]).AbsoluteUri)
    {
        Write-Host -ForegroundColor Red "Invalid value for the -AppDomain parameter"
        Exit
    }
}

# Requires an admin
if ($UseDeviceAuthentication)
{
    Connect-MgGraph -Scopes "Application.ReadWrite.All User.Read" -UseDeviceAuthentication -ErrorAction Stop
}
else
{
    Connect-MgGraph -Scopes "Application.ReadWrite.All User.Read" -ErrorAction Stop
}

# Find permission scope IDs
$UserReadScope = Find-MgGraphPermission -SearchString User.Read -PermissionType Delegated -ExactMatch -ErrorAction Stop
$CalendarsReadWriteScope = Find-MgGraphPermission -SearchString Calendars.ReadWrite -PermissionType Delegated -ExactMatch -ErrorAction Stop
$MailboxSettingsReadScope = Find-MgGraphPermission -SearchString MailboxSettings.Read -PermissionType Delegated -ExactMatch -ErrorAction Stop

# Create app registration
$CreateAppParams = @{
    DisplayName = $AppName
    Web = @{
        RedirectUris = @("https://" + $AppDomain + "/authcomplete")
    }
    RequiredResourceAccess = @{
        ResourceAppId = $GraphAppId
        ResourceAccess = @(
            @{
                Id = $UserReadScope.Id
                Type = "Scope"
            },
            @{
                Id = $CalendarsReadWriteScope.Id
                Type = "Scope"
            },
            @{
                Id = $MailboxSettingsReadScope.Id
                Type = "Scope"
            }
        )
    }
}

$appRegistration = New-MgApplication @CreateAppParams -ErrorAction Stop
Write-Host -ForegroundColor Cyan "App registration created with app ID" $appRegistration.AppId

# Expose an API (for Teams SSO)
$ScopeId = New-Guid
$ExposeApiParams = @{
    IdentifierUris = @(
        "api://" + $AppDomain + "/" + $appRegistration.AppId
    )
    Api = @{
        Oauth2PermissionScopes = @(
            @{
                Id = $ScopeId
                AdminConsentDisplayName = "Access the app as the user"
                AdminConsentDescription = "Allows Teams to call the app's web APIs as the current user"
                UserConsentDisplayName = "Access the app as you"
                UserConsentDescription = "Allows Teams to call the app's web APIs as you"
                IsEnabled = $true
                Type = "User"
                Value = "access_as_user"
            }
        )
        RequestedAccessTokenVersion = 2
    }
}

Update-MgApplication -ApplicationId $appRegistration.Id @ExposeApiParams -Debug -ErrorAction Stop
Write-Host -ForegroundColor Cyan "App registration updated to expose an API"

$AddPreAuthAppsParams = @{
    Api = @{
        PreAuthorizedApplications = @(
            @{
                AppId = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346"
                DelegatedPermissionIds = @(
                    $ScopeId
                )
            },
            @{
                AppId = "1fec8e78-bce4-4aaf-ab1b-5451cc387264"
                DelegatedPermissionIds = @(
                    $ScopeId
                )
            }
        )
    }
}

Update-MgApplication -ApplicationId $appRegistration.Id @AddPreAuthAppsParams -Debug -ErrorAction Stop
Write-Host -ForegroundColor Cyan "Teams clients added as preauthorized apps"

# Create corresponding service principal
New-MgServicePrincipal -AppId $appRegistration.AppId -ErrorAction SilentlyContinue `
-ErrorVariable SPError | Out-Null
if ($SPError)
{
    Write-Host -ForegroundColor Red "A service principal for the app could not be created."
    Write-Host -ForegroundColor Red $SPError
    Exit
}

Write-Host -ForegroundColor Cyan "Service principal created"

# Add client secret
$clientSecret = Add-MgApplicationPassword -ApplicationId $appRegistration.Id `
 -PasswordCredential @{ DisplayName = "Added by PowerShell" } -ErrorAction Stop

Write-Host
Write-Host -ForegroundColor Green "SUCCESS"
Write-Host -ForegroundColor Cyan -NoNewline "Client ID: "
Write-Host -ForegroundColor Yellow $appRegistration.AppId
Write-Host -ForegroundColor Cyan -NoNewline "Client secret: "
Write-Host -ForegroundColor Yellow $clientSecret.SecretText
Write-Host -ForegroundColor Cyan -NoNewline "Secret expires: "
Write-Host -ForegroundColor Yellow $clientSecret.EndDateTime

if ($StayConnected -eq $false)
{
  Disconnect-MgGraph
  Write-Host "Disconnected from Microsoft Graph"
}
else
{
  Write-Host
  Write-Host -ForegroundColor Yellow `
   "The connection to Microsoft Graph is still active. To disconnect, use Disconnect-MgGraph"
}
