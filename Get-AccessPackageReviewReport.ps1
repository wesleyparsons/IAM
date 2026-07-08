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

.PARAMETER OutputCsv
    Path to write the report to as CSV. Defaults to .\AccessPackageReviewReport.csv.

.NOTES
    Requires modules: Microsoft.Graph.Authentication, Microsoft.Graph.Identity.Governance,
    Microsoft.Graph.EntitlementManagement
    Requires delegated/app scopes: EntitlementManagement.Read.All, AccessReview.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputCsv,

    [string]$OutputCsv = '.\AccessPackageReviewReport.csv',

    [string[]]$RequiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.EntitlementManagement',
        'Microsoft.Graph.Identity.Governance'
    )
)

function Ensure-Modules {
    foreach ($m in $RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $m)) {
            Write-Host "Installing module $m ..." -ForegroundColor Yellow
            Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber
        }
        Import-Module -Name $m -ErrorAction Stop
    }
}

function Connect-ToGraph {
    $scopes = @('EntitlementManagement.Read.All', 'AccessReview.Read.All')
    $ctx = Get-MgContext
    if (-not $ctx -or (Compare-Object $ctx.Scopes $scopes -PassThru | Where-Object { $scopes -contains $_ }).Count -lt $scopes.Count) {
        Connect-MgGraph -Scopes $scopes | Out-Null
    }
    Write-Host "Connected as $((Get-MgContext).Account)" -ForegroundColor Green
}

function Select-AccessPackagesFromCsv {
    param([string]$Path)

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

    Write-Host "`nFetching access packages..." -ForegroundColor Cyan
    $allPackages = Get-MgEntitlementManagementAccessPackage -All

    $matched = New-Object System.Collections.Generic.List[object]
    foreach ($row in $rows) {
        $value = $row.$nameColumn
        if (-not $value) { continue }

        $hit = if ($nameColumn -eq 'Id') {
            $allPackages | Where-Object { $_.Id -eq $value }
        } else {
            $allPackages | Where-Object { $_.DisplayName -eq $value }
        }

        if ($hit) {
            $matched.Add($hit)
        } else {
            Write-Warning "No access package matched '$value' from CSV."
        }
    }

    return $matched
}

function Get-ReviewDefinitionsForPackage {
    param($AccessPackage)

    # Access package review definitions are ordinary accessReviews/definitions whose
    # scope.query embeds the access package's id (there's no direct "reviews for this
    # package" endpoint), so definitions are pulled once and matched by ID substring.
    # Do NOT use -ExpandProperty Instances here: for entitlement-management (access
    # package) reviews it hits a backend path that throws "BusinessFlow not found for
    # Id" 404s. Instances are fetched separately in Get-LatestInstance instead.
    $allDefinitions = Get-MgIdentityGovernanceAccessReviewDefinition -All

    return $allDefinitions | Where-Object {
        $scopeQuery = $_.Scope.AdditionalProperties['query']
        $scopeQuery -and $scopeQuery -like "*$($AccessPackage.Id)*"
    }
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
        Write-Warning "    v1.0 instance lookup failed for '$($Definition.DisplayName)' ($($Definition.Id)): $($_.Exception.Message). Retrying against beta..."
        try {
            $betaResult = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/beta/identityGovernance/accessReviews/definitions/$($Definition.Id)/instances?`$top=999" `
                -ErrorAction Stop
            $instances = $betaResult.value
        }
        catch {
            Write-Warning "    beta instance lookup also failed for '$($Definition.DisplayName)' ($($Definition.Id)): $($_.Exception.Message)"
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
        Write-Warning "    v1.0 decision lookup failed for '$($Definition.DisplayName)' instance $instanceId : $($_.Exception.Message). Retrying against beta..."
        try {
            $betaResult = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/beta/identityGovernance/accessReviews/definitions/$($Definition.Id)/instances/$instanceId/decisions?`$top=999" `
                -ErrorAction Stop
            $decisions = $betaResult.value
        }
        catch {
            Write-Warning "    beta decision lookup also failed for '$($Definition.DisplayName)' instance $instanceId : $($_.Exception.Message)"
            return @()
        }
    }

    foreach ($d in $decisions) {
        $reviewedBy = Get-Prop $d 'ReviewedBy'
        [PSCustomObject]@{
            AccessPackage = $AccessPackage.DisplayName
            ReviewName    = $Definition.DisplayName
            ReviewEnded   = Get-Prop $Instance 'EndDateTime'
            User          = Get-Prop (Get-Prop $d 'Principal') 'DisplayName'
            Reviewer      = if ($reviewedBy) { Get-Prop $reviewedBy 'DisplayName' } else { '(not reviewed)' }
            Outcome       = Get-Prop $d 'Decision'
            Justification = Get-Prop $d 'Justification'
        }
    }
}

# --- main ---

Ensure-Modules
Connect-ToGraph

$selectedPackages = Select-AccessPackagesFromCsv -Path $InputCsv
if (-not $selectedPackages) {
    Write-Warning "Nothing selected. Exiting."
    return
}

$report = New-Object System.Collections.Generic.List[object]

foreach ($pkg in $selectedPackages) {
    Write-Host "`nProcessing '$($pkg.DisplayName)'..." -ForegroundColor Cyan

    $definitions = Get-ReviewDefinitionsForPackage -AccessPackage $pkg
    if (-not $definitions) {
        Write-Warning "  No access review definitions found for this package."
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
}

if ($report.Count -eq 0) {
    Write-Warning "No review data to display."
    return
}

$report | Format-Table -AutoSize -Wrap

$report | Export-Csv -Path $OutputCsv -NoTypeInformation
Write-Host "`nExported to $OutputCsv" -ForegroundColor Green
