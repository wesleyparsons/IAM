<#
.SYNOPSIS
    Reports the outcome of the latest access review for a CSV-supplied list of access packages.

.DESCRIPTION
    Connects to Microsoft Graph, resolves the access packages listed in -InputCsv against
    Entitlement Management, finds the access review definition(s) scoped to each package,
    pulls the most recent review instance, and lists every decision: User, Reviewer,
    Outcome, Justification.

.PARAMETER InputCsv
    Path to a CSV listing the access packages to report on. The CSV needs one column,
    any of these header names: DisplayName, Name, AccessPackage, AccessPackageName, or Id.
    Not needed if -InputCsvBlobName is used instead; if both are given, the blob is
    downloaded to this path before being read.

.PARAMETER InputCsvBlobName
    Instead of a local -InputCsv, download the package list from a blob (e.g. so an
    Azure Automation runbook can pick up list edits without redeploying). Requires
    -StorageAccountName. Downloaded from -InputCsvContainerName using the same
    managed-identity/-UseConnectedAccount auth as -UploadToBlob.

.PARAMETER InputCsvContainerName
    Blob container to download the input CSV from. Defaults to 'access-review-input'.

.PARAMETER AccessPackagePrefix
    Alternative to -InputCsv/-InputCsvBlobName: one or more display-name prefixes (e.g.
    'PROD-', 'FIN-') to select access packages by, instead of reading a curated list.
    Every access package in the tenant is matched against every prefix (OR'd together).
    Mutually exclusive with -InputCsv/-InputCsvBlobName.

.PARAMETER OutputCsv
    Path to write the report to as CSV. Defaults to .\AccessPackageReviewReport.csv.

.PARAMETER UseManagedIdentity
    Connect to Microsoft Graph using the Azure Automation account's managed identity
    instead of an interactive/delegated sign-in. The identity's service principal needs
    the EntitlementManagement.Read.All and AccessReview.Read.All application permissions
    (admin-consented) - app-only auth has no delegated Entra role check on top of that.

.PARAMETER ManagedIdentityClientId
    Client ID of a user-assigned managed identity to use with -UseManagedIdentity. Omit
    to use the Automation account's system-assigned identity.

.PARAMETER UploadToBlob
    After writing -OutputCsv, also upload it to Azure Blob Storage for Power BI to read
    (via the Blob Storage connector's "combine files" feature, so history accumulates
    across runs). Requires -StorageAccountName. Auth uses the same managed identity as
    -UseManagedIdentity (or an interactive Connect-AzAccount if that switch is absent) -
    the identity needs the Storage Blob Data Contributor role on the container/account,
    no storage account key involved.

.PARAMETER StorageAccountName
    Storage account to upload the report to. Required when -UploadToBlob is set.

.PARAMETER StorageContainerName
    Blob container to upload into. Defaults to 'access-review-reports'.

.PARAMETER BlobName
    Blob name to upload as. Defaults to 'AccessPackageReviewReport_<timestamp>.csv' so
    each run lands as a new blob and Power BI's folder-combine picks up every run.

.NOTES
    Requires modules: Microsoft.Graph.Authentication, Microsoft.Graph.Identity.Governance,
    Microsoft.Graph.EntitlementManagement
    Requires delegated/app scopes: EntitlementManagement.Read.All, AccessReview.Read.All
    -UploadToBlob / -InputCsvBlobName additionally require: Az.Accounts, Az.Storage
#>

[CmdletBinding()]
param(
    [string]$InputCsv,

    [string]$InputCsvBlobName,

    [string]$InputCsvContainerName = 'access-review-input',

    [string[]]$AccessPackagePrefix,

    [string]$OutputCsv = '.\AccessPackageReviewReport.csv',

    [switch]$UseManagedIdentity,

    [string]$ManagedIdentityClientId,

    [switch]$UploadToBlob,

    [string]$StorageAccountName,

    [string]$StorageContainerName = 'access-review-reports',

    [string]$BlobName,

    [string[]]$RequiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.EntitlementManagement',
        'Microsoft.Graph.Identity.Governance'
    ),

    [string[]]$BlobRequiredModules = @(
        'Az.Accounts',
        'Az.Storage'
    )
)

function Ensure-Modules {
    param([string[]]$Modules = $RequiredModules)

    foreach ($m in $Modules) {
        if (-not (Get-Module -ListAvailable -Name $m)) {
            Write-Host "Installing module $m ..." -ForegroundColor Yellow
            Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber
        }
        Import-Module -Name $m -ErrorAction Stop
    }
}

function Connect-ToGraph {
    param(
        [switch]$UseManagedIdentity,
        [string]$ManagedIdentityClientId
    )

    if ($UseManagedIdentity) {
        Write-Host "`nConnecting to Microsoft Graph using managed identity..." -ForegroundColor Cyan
        if ($ManagedIdentityClientId) {
            Connect-MgGraph -Identity -ClientId $ManagedIdentityClientId | Out-Null
        } else {
            Connect-MgGraph -Identity | Out-Null
        }
        $ctx = Get-MgContext
        Write-Host "Connected via managed identity (client: $($ctx.ClientId))" -ForegroundColor Green
        # App-only auth has no delegated Entra role check - access is governed purely by
        # the identity's app role assignments (admin-consented application permissions).
        Write-Host "Note: this is app-only auth - make sure the identity's service principal has" -ForegroundColor DarkGray
        Write-Host "the EntitlementManagement.Read.All and AccessReview.Read.All application" -ForegroundColor DarkGray
        Write-Host "permissions admin-consented; there's no separate Entra role check for app-only calls." -ForegroundColor DarkGray
        return
    }

    $scopes = @('EntitlementManagement.Read.All', 'AccessReview.Read.All')
    $ctx = Get-MgContext
    if (-not $ctx -or (Compare-Object $ctx.Scopes $scopes -PassThru | Where-Object { $scopes -contains $_ }).Count -lt $scopes.Count) {
        Connect-MgGraph -Scopes $scopes | Out-Null
    }
    $ctx = Get-MgContext
    Write-Host "Connected as $($ctx.Account) | scopes: $($ctx.Scopes -join ', ')" -ForegroundColor Green
    # Reading access review instances/decisions is role-gated (Global Reader, Identity
    # Governance Administrator, Security Administrator, or the review's creator) on top
    # of the AccessReview.Read.All/EntitlementManagement.Read.All consent - having the
    # scope doesn't guarantee the signed-in account can read every review's results.
    Write-Host "Note: reading review results also requires an Entra role such as Identity" -ForegroundColor DarkGray
    Write-Host "Governance Administrator/Global Reader (or being the review's creator) -" -ForegroundColor DarkGray
    Write-Host "having the scope above doesn't by itself guarantee access to every review." -ForegroundColor DarkGray
}

function Connect-ToAzureStorage {
    param(
        [switch]$UseManagedIdentity,
        [string]$ManagedIdentityClientId
    )

    Write-Host "`nConnecting to Azure for blob upload..." -ForegroundColor Cyan
    if (-not (Get-AzContext)) {
        if ($UseManagedIdentity) {
            if ($ManagedIdentityClientId) {
                Connect-AzAccount -Identity -AccountId $ManagedIdentityClientId | Out-Null
            } else {
                Connect-AzAccount -Identity | Out-Null
            }
        } else {
            Connect-AzAccount | Out-Null
        }
    }
}

# Uses Azure AD/managed-identity auth against the blob data plane (-UseConnectedAccount)
# rather than a storage account key, so the only prerequisite is a Storage Blob Data
# Contributor role assignment for the signed-in identity on the container/account.
function Get-BlobStorageContext {
    param([string]$StorageAccountName)
    return New-AzStorageContext -StorageAccountName $StorageAccountName -UseConnectedAccount
}

function Publish-ReportToBlob {
    param(
        [string]$Path,
        [string]$StorageAccountName,
        [string]$ContainerName,
        [string]$BlobName
    )

    $ctx = Get-BlobStorageContext -StorageAccountName $StorageAccountName

    Write-Host "Uploading '$Path' to '$ContainerName/$BlobName' in storage account '$StorageAccountName'..." -ForegroundColor Cyan
    Set-AzStorageBlobContent -File $Path -Container $ContainerName -Blob $BlobName -Context $ctx -Force | Out-Null
    Write-Host "Uploaded. Point Power BI's Blob Storage connector at container '$ContainerName' and combine files to build history across runs." -ForegroundColor Green
}

function Get-InputCsvFromBlob {
    param(
        [string]$StorageAccountName,
        [string]$ContainerName,
        [string]$BlobName,
        [string]$DestinationPath
    )

    $ctx = Get-BlobStorageContext -StorageAccountName $StorageAccountName

    Write-Host "`nDownloading input CSV '$ContainerName/$BlobName' from storage account '$StorageAccountName'..." -ForegroundColor Cyan
    Get-AzStorageBlobContent -Container $ContainerName -Blob $BlobName -Destination $DestinationPath -Context $ctx -Force | Out-Null
    Write-Host "Downloaded to $DestinationPath" -ForegroundColor Green
}

# Pulls the full Graph error body (code/message/request-id) out of a failed Mg cmdlet
# call - $_.Exception.Message alone is often just a generic wrapper like "Response
# status code does not indicate success: NotFound (Not Found)."
function Get-GraphErrorDetail {
    param($ErrorRecord)

    $detail = $null
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        try {
            $parsed = $ErrorRecord.ErrorDetails.Message | ConvertFrom-Json -ErrorAction Stop
            if ($parsed.error) {
                $detail = "$($parsed.error.code): $($parsed.error.message) (request-id: $($parsed.error.innerError.'request-id'))"
            }
        }
        catch {
            $detail = $ErrorRecord.ErrorDetails.Message
        }
    }

    if (-not $detail) { $detail = $ErrorRecord.Exception.Message }
    return $detail
}

function Get-AllAccessPackages {
    Write-Host "`nFetching access packages..." -ForegroundColor Cyan
    return Get-MgEntitlementManagementAccessPackage -All
}

function Select-AccessPackagesFromCsv {
    param($AllPackages, [string]$Path)

    if (-not (Test-Path $Path)) {
        Write-Warning "CSV not found: $Path"
        return @()
    }

    $rows = Import-Csv -Path $Path
    if (-not $rows) {
        Write-Warning "CSV '$Path' has no rows."
        return @()
    }

    # accept whichever of these header names the caller used
    $nameColumn = ($rows[0].PSObject.Properties.Name |
        Where-Object { $_ -in @('DisplayName', 'Name', 'AccessPackage', 'AccessPackageName', 'Id') } |
        Select-Object -First 1)

    if (-not $nameColumn) {
        Write-Warning "CSV must have a column named one of: DisplayName, Name, AccessPackage, AccessPackageName, Id."
        return @()
    }

    $matched = New-Object System.Collections.Generic.List[object]
    foreach ($row in $rows) {
        $value = $row.$nameColumn
        if (-not $value) { continue }

        $hit = if ($nameColumn -eq 'Id') {
            $AllPackages | Where-Object { $_.Id -eq $value }
        } else {
            $AllPackages | Where-Object { $_.DisplayName -eq $value }
        }

        if ($hit) {
            $matched.Add($hit)
        } else {
            Write-Warning "No access package matched '$value' from CSV."
        }
    }

    return $matched
}

function Select-AccessPackagesByPrefix {
    param($AllPackages, [string[]]$Prefixes)

    $matched = New-Object System.Collections.Generic.List[object]
    foreach ($pkg in $AllPackages) {
        if ($Prefixes | Where-Object { $pkg.DisplayName -like "$_*" }) {
            $matched.Add($pkg)
        }
    }

    if (-not $matched) {
        Write-Warning "No access packages matched prefixes: $($Prefixes -join ', ')"
    } else {
        Write-Host "Matched $($matched.Count) access package(s) for prefixes: $($Prefixes -join ', ')" -ForegroundColor Green
    }

    return $matched
}

# Access review definitions are pulled once for the whole run (not per package) since
# there's no server-side "reviews for this access package" filter - every definition
# has to be scanned and matched client-side regardless of how many packages are asked for.
function Get-AllReviewDefinitions {
    Write-Host "`nFetching all access review definitions in the tenant..." -ForegroundColor Cyan
    return Get-MgIdentityGovernanceAccessReviewDefinition -All
}

function Get-ReviewDefinitionsForPackage {
    param($AllDefinitions, $AccessPackage)

    # Primary match: the access package's id literally appears in scope.query, e.g.
    # ".../accessPackageAssignments?$filter=(accessPackageId eq '{id}' ...)".
    $byScope = $AllDefinitions | Where-Object {
        $scopeQuery = $_.Scope.AdditionalProperties['query']
        $scopeQuery -and $scopeQuery -like "*$($AccessPackage.Id)*"
    }

    # Secondary match: definitions Entra auto-names after the access package but whose
    # scope query doesn't (or no longer) contain the id - e.g. a completed one-off review
    # from a since-edited/replaced assignment policy. Reported separately so a false
    # positive here is obvious rather than silently trusted like a scope-query match.
    $byName = $AllDefinitions | Where-Object {
        $_.DisplayName -and $_.DisplayName -like "*$($AccessPackage.DisplayName)*" -and
        ($byScope.Id -notcontains $_.Id)
    }

    foreach ($m in $byScope) {
        Write-Host "  Matched by scope: '$($m.DisplayName)' (Id: $($m.Id)) query: $($m.Scope.AdditionalProperties['query'])" -ForegroundColor DarkGray
    }
    foreach ($m in $byName) {
        Write-Host "  Matched by name only (verify this is really the same package): '$($m.DisplayName)' (Id: $($m.Id))" -ForegroundColor Yellow
    }

    return @($byScope) + @($byName)
}

function Show-AllDefinitionsDump {
    param($AllDefinitions, $AccessPackage)

    Write-Warning "  No results found for '$($AccessPackage.DisplayName)'. Dumping every access review definition in the tenant so you can spot the right one:"
    $AllDefinitions | ForEach-Object {
        [PSCustomObject]@{
            DisplayName = $_.DisplayName
            Id          = $_.Id
            ScopeQuery  = $_.Scope.AdditionalProperties['query']
        }
    } | Format-Table -AutoSize -Wrap | Out-String | Write-Host
}

function Get-LatestInstance {
    param($Definition)

    try {
        $instances = Get-MgIdentityGovernanceAccessReviewDefinitionInstance `
            -AccessReviewScheduleDefinitionId $Definition.Id -All -ErrorAction Stop
    }
    catch {
        # v1.0 sometimes can't resolve instances for entitlement-management (access
        # package) reviews ("BusinessFlow not found for Id"); the beta endpoint for
        # the same resource does. Fall back to a direct REST call before giving up.
        Write-Warning "    v1.0 instance lookup failed for '$($Definition.DisplayName)' ($($Definition.Id)): $(Get-GraphErrorDetail $_). Retrying against beta..."
        try {
            $betaResult = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/beta/identityGovernance/accessReviews/definitions/$($Definition.Id)/instances?`$top=999" `
                -ErrorAction Stop
            $instances = $betaResult.value
        }
        catch {
            Write-Warning "    beta instance lookup also failed for '$($Definition.DisplayName)' ($($Definition.Id)): $(Get-GraphErrorDetail $_)"
            return $null
        }
    }

    if (-not $instances) { return $null }

    $completed = $instances | Where-Object { $_.Status -in @('Completed', 'Applied') }
    $pool = if ($completed) { $completed } else { $instances }

    return $pool | Sort-Object EndDateTime -Descending | Select-Object -First 1
}

# Beta REST fallback returns plain (camelCase) hashtables instead of the typed
# SDK model, so property lookups have to work for both shapes.
function Get-Prop {
    param($Obj, [string]$Name)

    if ($null -eq $Obj) { return $null }
    if ($Obj -is [System.Collections.IDictionary]) { return $Obj[$Name] }

    $val = $Obj.$Name
    if ($null -ne $val) { return $val }

    # SDK's AdditionalProperties is a case-sensitive Dictionary<string,object> holding
    # Graph's original camelCase keys, so a PascalCase lookup needs a camelCase fallback.
    if ($Obj.AdditionalProperties) {
        if ($Obj.AdditionalProperties.ContainsKey($Name)) { return $Obj.AdditionalProperties[$Name] }
        $camelName = $Name.Substring(0, 1).ToLower() + $Name.Substring(1)
        if ($Obj.AdditionalProperties.ContainsKey($camelName)) { return $Obj.AdditionalProperties[$camelName] }
    }
    return $null
}

# EndDateTime comes back as a [datetime] from the SDK but as an ISO-8601 string from the
# beta REST fallback, and can be $null for an in-progress instance - normalize both.
function ConvertTo-NullableDateTime {
    param($Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value }
    try { return [datetime]$Value } catch { return $null }
}

function Get-DecisionRows {
    param($AccessPackage, $Definition, $Instance)

    $instanceId = Get-Prop $Instance 'Id'

    try {
        $decisions = Get-MgIdentityGovernanceAccessReviewDefinitionInstanceDecision `
            -AccessReviewScheduleDefinitionId $Definition.Id `
            -AccessReviewInstanceId $instanceId `
            -All -ErrorAction Stop
    }
    catch {
        Write-Warning "    v1.0 decision lookup failed for '$($Definition.DisplayName)' instance $instanceId : $(Get-GraphErrorDetail $_). Retrying against beta..."
        try {
            $betaResult = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/beta/identityGovernance/accessReviews/definitions/$($Definition.Id)/instances/$instanceId/decisions?`$top=999" `
                -ErrorAction Stop
            $decisions = $betaResult.value
        }
        catch {
            Write-Warning "    beta decision lookup also failed for '$($Definition.DisplayName)' instance $instanceId : $(Get-GraphErrorDetail $_)"
            return @()
        }
    }

    $reviewEndedRaw = Get-Prop $Instance 'EndDateTime'
    $reviewEndedDate = ConvertTo-NullableDateTime $reviewEndedRaw

    foreach ($d in $decisions) {
        $reviewedBy = Get-Prop $d 'ReviewedBy'
        [PSCustomObject]@{
            AccessPackage = $AccessPackage.DisplayName
            ReviewName    = $Definition.DisplayName
            ReviewEnded   = $reviewEndedRaw
            ReviewYear    = if ($reviewEndedDate) { $reviewEndedDate.Year } else { $null }
            ReviewQuarter = if ($reviewEndedDate) { [int][math]::Ceiling($reviewEndedDate.Month / 3.0) } else { $null }
            User          = Get-Prop (Get-Prop $d 'Principal') 'DisplayName'
            Reviewer      = if ($reviewedBy) { Get-Prop $reviewedBy 'DisplayName' } else { '(not reviewed)' }
            Outcome       = Get-Prop $d 'Decision'
            Justification = Get-Prop $d 'Justification'
        }
    }
}

# --- main ---

$usingCsv = [bool]($InputCsv -or $InputCsvBlobName)
$usingPrefix = [bool]($AccessPackagePrefix -and $AccessPackagePrefix.Count -gt 0)

if (-not $usingCsv -and -not $usingPrefix) {
    throw "Specify a package selection method: -InputCsv, -InputCsvBlobName, or -AccessPackagePrefix."
}
if ($usingCsv -and $usingPrefix) {
    throw "Specify only one package selection method: CSV (-InputCsv/-InputCsvBlobName) or -AccessPackagePrefix, not both."
}
if (($UploadToBlob -or $InputCsvBlobName) -and -not $StorageAccountName) {
    throw "-StorageAccountName is required when -UploadToBlob or -InputCsvBlobName is specified."
}

Ensure-Modules
Connect-ToGraph -UseManagedIdentity:$UseManagedIdentity -ManagedIdentityClientId $ManagedIdentityClientId

if ($InputCsvBlobName) {
    Ensure-Modules -Modules $BlobRequiredModules
    Connect-ToAzureStorage -UseManagedIdentity:$UseManagedIdentity -ManagedIdentityClientId $ManagedIdentityClientId

    $downloadPath = if ($InputCsv) { $InputCsv } else { Join-Path ([System.IO.Path]::GetTempPath()) 'AccessPackageReviewInput.csv' }
    Get-InputCsvFromBlob -StorageAccountName $StorageAccountName -ContainerName $InputCsvContainerName -BlobName $InputCsvBlobName -DestinationPath $downloadPath
    $InputCsv = $downloadPath
}

$allPackages = Get-AllAccessPackages

$selectedPackages = if ($usingPrefix) {
    Select-AccessPackagesByPrefix -AllPackages $allPackages -Prefixes $AccessPackagePrefix
} else {
    Select-AccessPackagesFromCsv -AllPackages $allPackages -Path $InputCsv
}
if (-not $selectedPackages) {
    Write-Warning "Nothing selected. Exiting."
    return
}

$allDefinitions = Get-AllReviewDefinitions

$report = New-Object System.Collections.Generic.List[object]

foreach ($pkg in $selectedPackages) {
    Write-Host "`nProcessing '$($pkg.DisplayName)'..." -ForegroundColor Cyan
    $rowCountBefore = $report.Count

    $definitions = Get-ReviewDefinitionsForPackage -AllDefinitions $allDefinitions -AccessPackage $pkg
    if (-not $definitions) {
        Write-Warning "  No access review definitions found for this package."
        Show-AllDefinitionsDump -AllDefinitions $allDefinitions -AccessPackage $pkg
        continue
    }

    foreach ($def in $definitions) {
        $latest = Get-LatestInstance -Definition $def
        if (-not $latest) {
            Write-Warning "  No review instances found for definition '$($def.DisplayName)'."
            continue
        }

        $rows = Get-DecisionRows -AccessPackage $pkg -Definition $def -Instance $latest
        foreach ($r in $rows) { $report.Add($r) }
    }

    if ($report.Count -eq $rowCountBefore) {
        Show-AllDefinitionsDump -AllDefinitions $allDefinitions -AccessPackage $pkg
    }
}

if ($report.Count -eq 0) {
    Write-Warning "No review data to display."
    return
}

$report | Format-Table -AutoSize -Wrap

$report | Export-Csv -Path $OutputCsv -NoTypeInformation
Write-Host "`nExported to $OutputCsv" -ForegroundColor Green

if ($UploadToBlob) {
    Ensure-Modules -Modules $BlobRequiredModules
    Connect-ToAzureStorage -UseManagedIdentity:$UseManagedIdentity -ManagedIdentityClientId $ManagedIdentityClientId

    $blobName = if ($BlobName) { $BlobName } else { "AccessPackageReviewReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" }
    Publish-ReportToBlob -Path $OutputCsv -StorageAccountName $StorageAccountName -ContainerName $StorageContainerName -BlobName $blobName
}
