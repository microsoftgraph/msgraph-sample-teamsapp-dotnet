# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT license.

param(
  [Parameter(Mandatory=$true,
  HelpMessage="The app ID of the app registration")]
  [String]
  $AppId,

  [Parameter(Mandatory=$true,
  HelpMessage="The new domain to set on the app registration")]
  [String[]]
  $AppDomain,

  [Parameter(Mandatory=$false)]
  [Switch]
  $UseDeviceAuthentication = $false,

  [Parameter(Mandatory=$false)]
  [Switch]
  $StayConnected = $false
)

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

# Get the application
$appRegistration = Get-MgApplication -Filter ("appId eq '" + $AppId + "'") -ErrorAction Stop

$UpdateAppParams = @{
    Web = @{
        RedirectUris = @("https://" + $AppDomain + "/authcomplete")
    }
    IdentifierUris = @(
        "api://" + $AppDomain + "/" + $appRegistration.AppId
    )
}

Update-MgApplication -ApplicationId $appRegistration.Id @UpdateAppParams -ErrorAction Stop

Write-Host
Write-Host -ForegroundColor Green "SUCCESS"
