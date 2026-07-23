# Get-AccessPackageReviewReport.ps1

Reports the outcome of the latest access review for a set of Entitlement Management
access packages. For each package it finds the matching access review definition(s),
pulls the most recent completed instance, and lists every decision: `User`, `Reviewer`,
`Outcome`, `Justification` — plus `ReviewYear` and `ReviewQuarter` (derived from the
review's end date) so the output can be filtered by calendar year/quarter downstream.

## How package selection works

Pick exactly one of these (mutually exclusive):

- **CSV** — `-InputCsv <path>` (local file) or `-InputCsvBlobName <name>` (downloaded
  from Blob Storage). The CSV needs one column named any of: `DisplayName`, `Name`,
  `AccessPackage`, `AccessPackageName`, or `Id`.
- **Prefix match** — `-AccessPackagePrefix <prefix1>,<prefix2>,...` selects every access
  package whose display name starts with any of the given prefixes, no CSV needed.

## How output works

- Always writes a local CSV: `-OutputCsv` (defaults to `.\AccessPackageReviewReport.csv`).
- Optionally also uploads that CSV to Blob Storage with `-UploadToBlob` (for Power BI to
  consume via its Blob Storage connector). Each run uploads under a new timestamped blob
  name by default, so Power BI's "combine files" feature accumulates history across runs
  rather than overwriting a single snapshot.

## Authentication

- Interactively (default): prompts an interactive `Connect-MgGraph` sign-in. If using
  Blob Storage input/output, also prompts an interactive `Connect-AzAccount` sign-in.
- In Azure Automation: pass `-UseManagedIdentity` to use the Automation account's
  system-assigned managed identity for both Graph and Storage auth (no prompts). Pass
  `-ManagedIdentityClientId <id>` instead if using a user-assigned identity.
  - The identity's service principal needs the Graph application permissions
    `EntitlementManagement.Read.All` and `AccessReview.Read.All` (admin-consented).
  - If using Blob Storage input/output, the identity also needs the **Storage Blob Data
    Contributor** role on the storage account/container (no storage account key used).

## Required modules

- Always: `Microsoft.Graph.Authentication`, `Microsoft.Graph.EntitlementManagement`,
  `Microsoft.Graph.Identity.Governance`
- Only if using Blob Storage input/output (`-UploadToBlob` or `-InputCsvBlobName`):
  `Az.Accounts`, `Az.Storage`

Modules are auto-installed (`Install-Module -Scope CurrentUser`) if missing.

## Examples

### 1. Interactive, CSV in / CSV out

```powershell
.\Get-AccessPackageReviewReport.ps1 `
    -InputCsv .\packages.csv `
    -OutputCsv .\AccessPackageReviewReport.csv
```

Prompts an interactive Graph sign-in, reads the package list from `packages.csv`, and
writes the report to `AccessPackageReviewReport.csv`.

### 2. Interactive, prefix match / CSV out

```powershell
.\Get-AccessPackageReviewReport.ps1 `
    -AccessPackagePrefix 'PROD-', 'FIN-' `
    -OutputCsv .\AccessPackageReviewReport.csv
```

No CSV needed — selects every access package whose display name starts with `PROD-` or
`FIN-`.

### 3. Azure Automation, CSV in / CSV+Blob out

```powershell
.\Get-AccessPackageReviewReport.ps1 `
    -UseManagedIdentity `
    -InputCsv .\packages.csv `
    -UploadToBlob `
    -StorageAccountName myiamstorage
```

Use this when the package list is small and stable enough to ship as a file alongside
the runbook (e.g. baked into the runbook's working directory). Authenticates via managed
identity, writes the local CSV, then uploads it to the `access-review-reports` container
(default `-StorageContainerName`) for Power BI to pick up.

### 4. Azure Automation, Blob in / Blob out (fully unattended)

```powershell
.\Get-AccessPackageReviewReport.ps1 `
    -UseManagedIdentity `
    -InputCsvBlobName packages.csv `
    -UploadToBlob `
    -StorageAccountName myiamstorage
```

The package list itself is edited by uploading a new blob to the
`access-review-input` container (default `-InputCsvContainerName`) — no runbook
redeploy needed to change the list. Same storage account and managed identity handle
both the input download and the output upload; the identity needs **Storage Blob Data
Contributor** on it.

### 5. Azure Automation, prefix match / Blob out

```powershell
.\Get-AccessPackageReviewReport.ps1 `
    -UseManagedIdentity `
    -AccessPackagePrefix 'PROD-', 'FIN-' `
    -UploadToBlob `
    -StorageAccountName myiamstorage
```

No CSV at all — packages are selected by prefix, report is uploaded to Blob Storage for
Power BI.

## Parameters reference

| Parameter | Purpose | Default |
|---|---|---|
| `-InputCsv` | Local path to the package list CSV | — |
| `-InputCsvBlobName` | Blob name to download the package list from | — |
| `-InputCsvContainerName` | Container to download the input CSV from | `access-review-input` |
| `-AccessPackagePrefix` | Display-name prefix(es) to select packages by | — |
| `-OutputCsv` | Local path to write the report to | `.\AccessPackageReviewReport.csv` |
| `-UseManagedIdentity` | Use managed identity instead of interactive sign-in | off |
| `-ManagedIdentityClientId` | Client ID of a user-assigned managed identity | system-assigned |
| `-UploadToBlob` | Also upload the report CSV to Blob Storage | off |
| `-StorageAccountName` | Storage account for blob input/output | — (required if either blob feature is used) |
| `-StorageContainerName` | Container to upload the report to | `access-review-reports` |
| `-BlobName` | Blob name for the uploaded report | `AccessPackageReviewReport_<timestamp>.csv` |

## Power BI

Point Power BI's **Blob Storage** connector at the `access-review-reports` container and
use "combine files" — each run's CSV lands as a separate timestamped blob, so combining
them builds up history that can be sliced by `ReviewYear`/`ReviewQuarter` in Power BI.
