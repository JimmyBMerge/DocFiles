<#
.SYNOPSIS
    Azure Automation runbook script to extract the structure of a SharePoint Online site and persist it to Azure Blob Storage with blob versioning.

.DESCRIPTION
    Uses an Azure AD app registration (client secret credential) to authenticate against Microsoft Graph and Azure Storage. Enumerates lists, document library folder structure, and (when applicable) Microsoft Teams channel tabs associated with the site collection's Microsoft 365 group. The resulting snapshot is serialized as JSON and uploaded to a blob. When the storage account has blob versioning enabled, each snapshot upload will create a new version automatically.


.PARAMETER SiteId
    SharePoint site identifier. Accepts any of the following forms:
    - Full Graph composite identifier (e.g.
      contoso.sharepoint.com,39999de1-97a2-40e5-9c1d-95288ba1d0fa,fad53fcc-ac9a-4060-87b9-f9f0e69c0dcb)
    - Site GUID only (requires SiteHostname and SitePath to expand to the composite form)
    - Blank/omitted (the script resolves the identifier using SiteHostname and SitePath)
    When not provided, the script attempts to read the value from an Automation variable named
    `SiteId`.

.PARAMETER SiteHostname
    SharePoint Online host name (e.g. contoso.sharepoint.com). When the site identifier is not
    supplied as a runbook parameter, the script attempts to read the host name from an Automation
    variable named `SiteHostname`.

.PARAMETER SitePath
    Site-relative path (e.g. sites/ProjectX). When the site identifier is not supplied, the script
    attempts to read the value from an Automation variable named `SitePath`.

.PARAMETER TenantId
    Azure AD tenant ID associated with the app registration. When omitted, the script looks for an
    Automation variable named `TenantId`.

.PARAMETER ClientId
    Client ID (application ID) of the Azure AD app registration. If this parameter is not
    specified the runbook uses the username from the `SharePointAppRegistration` Automation
    credential.

.PARAMETER ClientSecret
    Secure string containing the client secret associated with the app registration. When not
    supplied, the password from the `SharePointAppRegistration` Automation credential is used.

.PARAMETER StorageAccountName
    Target storage account name where the snapshot should be stored. When omitted, the script
    attempts to read the value from an Automation variable named `StorageAccountName`.

.PARAMETER StorageResourceGroup
    Resource group containing the storage account. When omitted, the script attempts to read the
    value from an Automation variable named `StorageResourceGroup`.

.PARAMETER StorageContainerName
    Blob container name where snapshots should be written. When omitted, the script attempts to
    read the value from an Automation variable named `StorageContainerName`.

.PARAMETER OutputBlobPrefix
    Optional blob name prefix. Timestamp is appended automatically. When omitted, the script looks
    for an Automation variable named `OutputBlobPrefix` and, if that is not found, defaults to
    `site-structure`.

.EXAMPLE
    .\Extract-SharePointStructure.ps1

    Creates a JSON snapshot of the SharePoint site's structure and stores it in the specified
    container. All required values are resolved from Automation assets (variables and the
    `SharePointAppRegistration` credential). Provide explicit parameter values to override the
    Automation defaults when testing locally.

.NOTES
    When importing the script into an Azure Automation PowerShell runbook, the parameters in the `param` block will surface as runbook input fields. You can provide values by:
    - Supplying them interactively when you click **Start** on the runbook blade in the Azure portal.
    - Linking a schedule/webhook and entering default parameter values in the runbook's **Parameters and run settings** pane.
    - Starting the runbook through PowerShell or REST. For example:
        ```powershell
        Start-AzAutomationRunbook -AutomationAccountName "ContosoAutomation" -ResourceGroupName "rg-automation" `
            -Name "Extract-SharePointStructure" `
            -Parameters @{ OutputBlobPrefix = 'projectx-site' }
        ```
    Configure Automation variables named `SiteId` (or `SiteHostname` and `SitePath`), `TenantId`,
    `StorageAccountName`, `StorageResourceGroup`, `StorageContainerName`, and (optionally)
    `OutputBlobPrefix`. Store the
    app registration's client ID and secret in the `SharePointAppRegistration` Automation credential.
    When running the script outside Automation, you can still provide parameters explicitly.
#>

param(
    [string]$SiteId,
    [string]$SiteHostname,
    [string]$SitePath,
    [string]$TenantId,
    [string]$ClientId,
    [System.Security.SecureString]$ClientSecret,
    [string]$StorageAccountName,
    [string]$StorageResourceGroup,
    [string]$StorageContainerName,
    [string]$OutputBlobPrefix
)

function Resolve-RunbookValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowNull()][object]$SuppliedValue,
        [switch]$Optional,
        [AllowNull()][object]$DefaultValue
    )

    $suppliedString = if ($SuppliedValue -is [string]) { $SuppliedValue.Trim() } else { $null }

    if ($PSBoundParameters.ContainsKey($Name)) {
        if ($SuppliedValue -is [string]) {
            if (-not [string]::IsNullOrWhiteSpace($suppliedString)) {
                return $SuppliedValue
            }
        }
        elseif ($null -ne $SuppliedValue) {
            return $SuppliedValue
        }
    }

    if ($null -ne $SuppliedValue -and -not ($SuppliedValue -is [string] -and [string]::IsNullOrWhiteSpace($suppliedString))) {
        return $SuppliedValue
    }

    $automationValue = $null
    if (Get-Command -Name Get-AutomationVariable -ErrorAction SilentlyContinue) {
        try {
            $automationValue = Get-AutomationVariable -Name $Name -ErrorAction Stop
        }
        catch {
            $automationValue = $null
        }
    }

    if ($null -ne $automationValue -and -not ($automationValue -is [string] -and [string]::IsNullOrWhiteSpace($automationValue))) {
        Write-Verbose "Resolved '$Name' from Automation variable."
        return $automationValue
    }

    if ($null -ne $DefaultValue) {
        return $DefaultValue
    }

    if ($Optional) {
        return $SuppliedValue
    }

    throw "Configuration value for '$Name' was not provided. Supply it as a parameter or create an Automation variable named '$Name'."
}

# region Authentication helpers
function ConvertFrom-SecureStringToPlainText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Security.SecureString]$SecureString
    )

    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function Get-GraphToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$ClientId,
        [Parameter(Mandatory = $true)][System.Security.SecureString]$ClientSecret
    )

    $secretPlainText = ConvertFrom-SecureStringToPlainText -SecureString $ClientSecret

    $body = @{
        client_id     = $ClientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $secretPlainText
        grant_type    = "client_credentials"
    }

    try {
        $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
                                            -Method Post `
                                            -Body $body `
                                            -ContentType "application/x-www-form-urlencoded"
        return $tokenResponse.access_token
    }
    finally {
        $secretPlainText = $null
    }
}

function Invoke-GraphGet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [hashtable]$Headers,
        [string]$GraphToken
    )

    if (-not $Headers) {
        $Headers = @{}
    }

    if (-not $Headers.ContainsKey('Authorization')) {
        if (-not $GraphToken) {
            throw "Graph token was not provided for request to $Uri."
        }
        $Headers['Authorization'] = "Bearer $GraphToken"
    }

    $Headers['ConsistencyLevel'] = 'eventual'
    Invoke-RestMethod -Method Get -Uri $Uri -Headers $Headers
}

function Get-GraphCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$GraphToken,
        [hashtable]$Headers
    )

    $results = @()
    $nextLink = $Uri

    while ($nextLink) {
        $response = Invoke-GraphGet -Uri $nextLink -Headers $Headers -GraphToken $GraphToken
        if ($response.value) {
            $results += $response.value
        }
        $nextLink = $response.'@odata.nextLink'
    }

    return ,$results
}
# endregion

# region SharePoint structure discovery
function Get-SiteLists {
    param(
        [Parameter(Mandatory = $true)][string]$SiteId,
        [Parameter(Mandatory = $true)][string]$GraphToken
    )

    $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists?`$select=id,displayName,list,webUrl,createdDateTime,lastModifiedDateTime"
    Get-GraphCollection -Uri $uri -GraphToken $GraphToken
}

function Normalize-SiteId {
    [CmdletBinding()]
    param(
        [string]$SiteId,
        [string]$SiteHostname,
        [string]$SitePath,
        [string]$GraphToken
    )

    $hasHostnamePath = -not [string]::IsNullOrWhiteSpace($SiteHostname) -and -not [string]::IsNullOrWhiteSpace($SitePath)

    if (-not [string]::IsNullOrWhiteSpace($SiteId)) {
        # Case 1: Composite SiteId (hostname,guid,guid)
        if ($SiteId -match ',') {
            $parts = $SiteId.Split(',')
            if ($parts.Count -ne 3) {
                throw "Invalid SiteId format: '$SiteId'. Expected format is 'hostname,siteCollectionId,siteId'."
            }

            $hostname, $collectionId, $siteGuid = $parts

            if ([string]::IsNullOrWhiteSpace($hostname)) {
                throw "Invalid SiteId: missing hostname in '$SiteId'."
            }

            if ($collectionId -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                throw "Invalid SiteId: siteCollectionId part '$collectionId' is not a valid GUID."
            }

            if ($siteGuid -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                throw "Invalid SiteId: siteId part '$siteGuid' is not a valid GUID."
            }

            Write-Output "Using supplied composite SiteId '$SiteId'."
            return $SiteId
        }

        # Case 2: GUID only
        if ($SiteId -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
            if (-not $hasHostnamePath) {
                throw "SiteId was provided as GUID only. Please also supply SiteHostname and SitePath to resolve full Graph site ID."
            }

            Write-Output "SiteId '$SiteId' detected as GUID. Resolving composite identifier using hostname/path..."
            $uri = "https://graph.microsoft.com/v1.0/sites/$($SiteHostname):/$($SitePath)"
            $site = Invoke-GraphGet -Uri $uri -GraphToken $GraphToken
            if (-not $site -or -not $site.id) {
                throw "Failed to resolve site using hostname '$SiteHostname' and path '$SitePath'."
            }

            Write-Output "Resolved composite SiteId '$($site.id)' from GUID '$SiteId'."
            return $site.id
        }

        # Case 3: Invalid format
        Write-Output "SiteId '$SiteId' not recognized as composite or GUID. Falling back to hostname/path..."
    }

    # Case 4: No SiteId or unrecognized value
    if (-not $hasHostnamePath) {
        throw "Either SiteId (composite or GUID) or both SiteHostname and SitePath must be provided."
    }

    Write-Output "Resolving SiteId using hostname/path $SiteHostname/$SitePath..."
    $fallbackUri = "https://graph.microsoft.com/v1.0/sites/$($SiteHostname):/$($SitePath)"
    $fallbackSite = Invoke-GraphGet -Uri $fallbackUri -GraphToken $GraphToken
    if (-not $fallbackSite -or -not $fallbackSite.id) {
        throw "Failed to resolve site using hostname '$SiteHostname' and path '$SitePath'."
    }

    Write-Output "Resolved composite SiteId '$($fallbackSite.id)' using hostname/path."
    return $fallbackSite.id
}

function Get-DriveStructure {
    param(
        [Parameter(Mandatory = $true)][string]$DriveId,
        [Parameter(Mandatory = $true)][string]$GraphToken,
        [string]$ParentPath = ''
    )

    $normalizedPath = if ($ParentPath) { ($ParentPath -replace '\\', '/') } else { $null }
    $encodedPath = if ($normalizedPath) {
        ($([System.Uri]::EscapeDataString($normalizedPath)) -replace '%2F', '/')
    } else {
        $null
    }

    $driveSegment = if ($encodedPath) { "/root:/$($encodedPath):/children" } else { "/root/children" }
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId$driveSegment?`$select=id,name,folder,webUrl"
    $children = Get-GraphCollection -Uri $uri -GraphToken $GraphToken

    foreach ($child in $children) {
        $item = [ordered]@{
            Id      = $child.id
            Name    = $child.name
            Path    = if ($ParentPath) { "$ParentPath/$($child.name)" } else { $child.name }
            WebUrl  = $child.webUrl
            Type    = if ($child.folder) { 'Folder' } else { 'File' }
            Children = @()
        }

        if ($child.folder) {
            $item.Children = Get-DriveStructure -DriveId $DriveId -GraphToken $GraphToken -ParentPath $item.Path
        }

        $item
    }
}

function Get-TeamTabs {
    param(
        [Parameter(Mandatory = $true)][string]$GroupId,
        [Parameter(Mandatory = $true)][string]$GraphToken
    )

    $headers = @{ Authorization = "Bearer $GraphToken" }
    $channelsUri = "https://graph.microsoft.com/v1.0/teams/$GroupId/channels?`$select=id,displayName"
    $channels = Get-GraphCollection -Uri $channelsUri -Headers $headers -GraphToken $GraphToken

    foreach ($channel in $channels) {
        $tabsUri = "https://graph.microsoft.com/v1.0/teams/$GroupId/channels/$($channel.id)/tabs?`$select=id,displayName,configuration&`$expand=teamsApp"
        $tabs = Get-GraphCollection -Uri $tabsUri -Headers $headers -GraphToken $GraphToken

        [ordered]@{
            ChannelId   = $channel.id
            ChannelName = $channel.displayName
            Tabs        = $tabs | ForEach-Object {
                [ordered]@{
                    TabId        = $_.id
                    DisplayName  = $_.displayName
                    AppId        = if ($_.teamsApp) { $_.teamsApp.id } else { $null }
                    Configuration = $_.configuration
                }
            }
        }
    }
}
# endregion

# region Storage helpers
function Get-StorageContext {
    param(
        [Parameter(Mandatory = $true)][string]$ResourceGroupName,
        [Parameter(Mandatory = $true)][string]$AccountName
    )

    $account = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $AccountName
    if (-not $account) {
        throw "Storage account '$AccountName' in resource group '$ResourceGroupName' was not found."
    }

    New-AzStorageContext -StorageAccountName $AccountName -UseConnectedAccount
}

function Ensure-Container {
    param(
        [Parameter(Mandatory = $true)]$Context,
        [Parameter(Mandatory = $true)][string]$ContainerName
    )

    $container = Get-AzStorageContainer -Context $Context -Name $ContainerName -ErrorAction SilentlyContinue
    if (-not $container) {
        Write-Verbose "Creating container '$ContainerName'."
        New-AzStorageContainer -Context $Context -Name $ContainerName -Permission Off | Out-Null
    }
}
# endregion

$SiteId = Resolve-RunbookValue -Name 'SiteId' -SuppliedValue $SiteId -Optional
$SiteHostname = Resolve-RunbookValue -Name 'SiteHostname' -SuppliedValue $SiteHostname -Optional
$SitePath = Resolve-RunbookValue -Name 'SitePath' -SuppliedValue $SitePath -Optional
$TenantId = Resolve-RunbookValue -Name 'TenantId' -SuppliedValue $TenantId
$ClientId = Resolve-RunbookValue -Name 'ClientId' -SuppliedValue $ClientId -Optional
$ClientSecret = Resolve-RunbookValue -Name 'ClientSecret' -SuppliedValue $ClientSecret -Optional
$StorageAccountName = Resolve-RunbookValue -Name 'StorageAccountName' -SuppliedValue $StorageAccountName
$StorageResourceGroup = Resolve-RunbookValue -Name 'StorageResourceGroup' -SuppliedValue $StorageResourceGroup
$StorageContainerName = Resolve-RunbookValue -Name 'StorageContainerName' -SuppliedValue $StorageContainerName
$OutputBlobPrefix = Resolve-RunbookValue -Name 'OutputBlobPrefix' -SuppliedValue $OutputBlobPrefix -Optional -DefaultValue 'site-structure'

$automationCredential = $null
if ((-not $ClientId) -or (-not $ClientSecret)) {
    if (Get-Command -Name Get-AutomationPSCredential -ErrorAction SilentlyContinue) {
        try {
            $automationCredential = Get-AutomationPSCredential -Name 'SharePointAppRegistration' -ErrorAction Stop
        }
        catch {
            $automationCredential = $null
        }
    }

    if (-not $ClientId -and $automationCredential) {
        $ClientId = $automationCredential.UserName
        Write-Verbose "Resolved 'ClientId' from Automation credential 'SharePointAppRegistration'."
    }

    if (-not $ClientSecret -and $automationCredential) {
        $ClientSecret = $automationCredential.Password
        Write-Verbose "Resolved 'ClientSecret' from Automation credential 'SharePointAppRegistration'."
    }
}

if ($ClientSecret -and $ClientSecret -isnot [System.Security.SecureString]) {
    if ($ClientSecret -is [string]) {
        $ClientSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
    }
    else {
        throw "ClientSecret must be a SecureString."
    }
}

if (-not $ClientId -or -not $ClientSecret) {
    throw "Client credentials were not provided. Supply ClientId/ClientSecret parameters, create Automation variables with those names, or configure the 'SharePointAppRegistration' Automation credential."
}

try {
    Write-Output "Connecting to Azure with app registration credentials..."
    $servicePrincipalCredential = New-Object System.Management.Automation.PSCredential($ClientId, $ClientSecret)
    Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $servicePrincipalCredential | Out-Null

    Write-Output "Requesting Microsoft Graph token (client credentials flow)..."
    $graphToken = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    if (-not $graphToken) {
        throw "Failed to obtain Microsoft Graph token using the provided app registration."
    }
    Write-Output "✅ Successfully obtained Graph token."

    $siteId = Normalize-SiteId -SiteId $SiteId -SiteHostname $SiteHostname -SitePath $SitePath -GraphToken $graphToken
    Write-Output "Resolving SharePoint site details using composite id '$siteId'..."
    $site = Invoke-GraphGet -Uri "https://graph.microsoft.com/v1.0/sites/$siteId" -GraphToken $graphToken
    if (-not $site) {
        throw "Unable to resolve site using identifier '$siteId'."
    }

    Write-Output "✅ Connected to site '$($site.displayName)' ($($site.webUrl))."
    if ($site.id) {
        Write-Output "Graph-reported site id: $($site.id)"
    }

    Write-Output "Enumerating lists for site id $siteId..."
    $lists = Get-SiteLists -SiteId $siteId -GraphToken $graphToken

    Write-Output "Enumerating drives and folder structure..."
    $drivesUri = "https://graph.microsoft.com/v1.0/sites/$siteId/drives?`$select=id,name,webUrl"
    $driveItems = Get-GraphCollection -Uri $drivesUri -GraphToken $graphToken
    $drives = $driveItems | ForEach-Object {
        [ordered]@{
            DriveId   = $_.id
            Name      = $_.name
            WebUrl    = $_.webUrl
            Structure = Get-DriveStructure -DriveId $_.id -GraphToken $graphToken
        }
    }

    $teams = $null
    if ($site.siteCollection -and $site.siteCollection.groupId) {
        Write-Output "Site is connected to Microsoft 365 group $($site.siteCollection.groupId). Enumerating Teams tabs..."
        $teams = Get-TeamTabs -GroupId $site.siteCollection.groupId -GraphToken $graphToken
    }

    $snapshot = [ordered]@{
        RetrievedOnUtc = (Get-Date).ToUniversalTime().ToString('o')
        Site           = $site | Select-Object id, displayName, webUrl, siteCollection
        Lists          = $lists
        Drives         = $drives
        Teams          = $teams
    }

    $json = $snapshot | ConvertTo-Json -Depth 15

    $context = Get-StorageContext -ResourceGroupName $StorageResourceGroup -AccountName $StorageAccountName
    Ensure-Container -Context $context -ContainerName $StorageContainerName

    $timestamp = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
    $blobName = "${OutputBlobPrefix}-${timestamp}.json"
    $tempFile = Join-Path -Path $env:TEMP -ChildPath $blobName
    $json | Out-File -FilePath $tempFile -Encoding UTF8

    Write-Output "Uploading snapshot to blob $blobName..."
    Set-AzStorageBlobContent -Context $context -File $tempFile -Container $StorageContainerName -Blob $blobName -Force | Out-Null
    Remove-Item $tempFile -ErrorAction SilentlyContinue

    Write-Output "Snapshot upload complete. Ensure blob versioning is enabled on the storage account to maintain history."
}
catch {
    Write-Error $_
    throw
}
