# EnvironmentDataLoader

PowerShell toolset for importing and migrating data in **Dynamics 365 Finance & Operations** via the Data Management Framework (DMF) API.

---

## Contents

- [Requirements](#requirements)
- [Repository structure](#repository-structure)
- [Authentication](#authentication)
- [Scripts](#scripts)
  - [Invoke-BaselineImport.ps1](#invoke-baselineimportps1)
  - [Invoke-TemplateExport.ps1](#invoke-templateexportps1)
  - [Invoke-ProjectExport.ps1](#invoke-projectexportps1)
  - [Expand-ExportedPackages.ps1](#expand-exportedpackagesps1)
  - [Invoke-PackageUpload.ps1](#invoke-packageuploadps1)
- [Full migration pipeline](#full-migration-pipeline)
- [Entity execution ordering](#entity-execution-ordering)
- [Library modules](#library-modules)

---

## Requirements

- PowerShell 5.1 or later
- Network access to the D365 F&O environment and Azure Blob Storage
- A Microsoft Entra (Azure AD) account with permissions to sign in to D365
- The **Data Management** workspace must be accessible in the target environment

---

## Repository structure

```
EnvironmentDataLoader/
├── Invoke-BaselineImport.ps1     Import local packages (xlsx + Manifest) into D365
├── Invoke-TemplateExport.ps1     Export templates to local .zip files (uses template directly)
├── Invoke-ProjectExport.ps1      Build a DMF project from template lines, then export
├── Expand-ExportedPackages.ps1   Extract downloaded .zip files into review folders
├── Invoke-PackageUpload.ps1      Upload pre-built .zip files into D365
├── resources/
│   └── 010 - System Setup/       Example baseline data package
│       ├── Manifest.xml
│       ├── PackageHeader.xml
│       └── *.xlsx
└── lib/
    ├── DmfOutput.ps1             Console output and transcript helpers
    ├── DmfRequest.ps1            REST client with automatic retry
    ├── DmfPackage.ps1            Package discovery and entity ordering
    └── DmfZip.ps1                DMF zip inspection helpers
```

---

## Authentication

All scripts that connect to D365 use the **Microsoft Entra device code flow**.  When authentication is required, the script prints a short code and a URL:

```
To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code XXXXXXXXX to authenticate.
```

Open the URL in any browser, enter the code, and sign in with your D365 credentials.  The token is obtained once per run and reused for all API calls in that session.

**Client ID used:** `1950a258-227b-4e31-a9cf-717495945fc2` (the public Azure CLI application — no app registration required).

---

## Scripts

### Invoke-BaselineImport.ps1

Discovers package folders under a local directory, builds a DMF zip from each one, uploads it to Azure Blob Storage, and imports it into D365 via `ImportFromPackage`.

Each package folder must contain:
- `Manifest.xml` — entity definitions and file mappings
- One or more `.xlsx` files — one per entity
- *(optional)* `ordering.json` — per-package entity execution overrides

The manifest is rebuilt before upload with optimised `ExecutionUnit / LevelInExecutionUnit / SequenceInLevel` values so that independent entity chains can run in parallel inside D365 DMF.

#### Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-EnvironmentUrl` | Yes | — | D365 base URL, e.g. `https://contoso.operations.dynamics.com` |
| `-TenantId` | Yes | — | Entra tenant ID or domain, e.g. `contoso.onmicrosoft.com` |
| `-LegalEntityId` | Yes | — | D365 company to import into, e.g. `DAT` |
| `-PackageName` | No | — | Import exactly this folder without showing the menu |
| `-ResourcesPath` | No | `./resources` | Root directory containing package subfolders |
| `-OutputPath` | No | `$env:TEMP` | Directory for the temporary zip files |
| `-LogPath` | No | auto | Transcript log path; pass `''` to suppress |
| `-PollIntervalSeconds` | No | `30` | Status check interval (5–300) |
| `-TimeoutMinutes` | No | `60` | Per-package polling timeout (1–480) |
| `-MaxRetries` | No | `3` | Retry limit for transient REST failures (0–10) |
| `-Force` | No | off | Skip the confirmation prompt |
| `-WhatIf` | No | off | Validate without making any API calls |
| `-NoOverwrite` | No | off | Preserve existing D365 records |
| `-KeepZip` | No | off | Keep the generated zip after upload |
| `-PassThru` | No | off | Emit result objects: `Package, Status, ExecutionId, Elapsed` |

#### Examples

```powershell
# Interactive: discover all packages and select which to import
.\Invoke-BaselineImport.ps1 `
    -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'DAT'
```

```powershell
# Non-interactive: import one package, skip confirmation, write a log
.\Invoke-BaselineImport.ps1 `
    -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'DAT' `
    -PackageName    '010 - System Setup' `
    -Force `
    -LogPath        'C:\Logs\import.log'
```

```powershell
# Dry-run: validate all packages without touching D365
.\Invoke-BaselineImport.ps1 `
    -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'DAT' `
    -WhatIf
```

```powershell
# Import from a custom directory into a non-default company
.\Invoke-BaselineImport.ps1 `
    -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'USMF' `
    -ResourcesPath  'C:\DMF\Packages' `
    -NoOverwrite
```

---

### Invoke-TemplateExport.ps1

Reads the list of DMF definition-group templates from a D365 environment, lets you select which ones to export, submits `ExportToPackage` jobs, polls for completion, and downloads the resulting zip files.

#### Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-EnvironmentUrl` | Yes | — | D365 base URL |
| `-TenantId` | Yes | — | Entra tenant ID or domain |
| `-LegalEntityId` | Yes | — | D365 company to export from |
| `-TemplateName` | No | — | Export exactly this template (DefinitionGroupId) without showing the menu |
| `-DownloadPath` | No | `$env:TEMP` | Directory to save downloaded zip files; pass `''` to get the URL only |
| `-LogPath` | No | auto | Transcript log path; pass `''` to suppress |
| `-PollIntervalSeconds` | No | `30` | Status check interval (5–300) |
| `-TimeoutMinutes` | No | `60` | Per-template polling timeout (1–480) |
| `-MaxRetries` | No | `3` | Retry limit for transient REST failures (0–10) |
| `-Force` | No | off | Skip the confirmation prompt |
| `-WhatIf` | No | off | With `-TemplateName`: no API calls at all. Without: fetches the template list (read-only) for the menu, then exits without exporting |
| `-PassThru` | No | off | Emit result objects: `Template, TemplateId, Status, ExecutionId, DownloadUrl, DownloadedTo, Elapsed` |

#### Examples

```powershell
# Interactive: list all templates in the environment and choose which to export
.\Invoke-TemplateExport.ps1 `
    -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'DAT'
```

```powershell
# Export a specific template and download the zip to C:\DMF\Downloads
.\Invoke-TemplateExport.ps1 `
    -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'DAT' `
    -TemplateName   'SystemSetupExport' `
    -DownloadPath   'C:\DMF\Downloads' `
    -Force
```

```powershell
# Export all templates, download every zip, capture results for further processing
$exports = .\Invoke-TemplateExport.ps1 `
    -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'DAT' `
    -DownloadPath   'C:\DMF\Downloads' `
    -Force `
    -PassThru

$exports | Where-Object Status -eq 'Succeeded' | Select-Object TemplateId, DownloadedTo
```

```powershell
# WhatIf with a named template — zero API calls, no authentication needed
.\Invoke-TemplateExport.ps1 `
    -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'DAT' `
    -TemplateName   'SystemSetupExport' `
    -WhatIf
```

---

### Invoke-ProjectExport.ps1

Builds a dedicated DMF export project from a template's entity lines, then exports and downloads it.  Unlike `Invoke-TemplateExport.ps1` — which runs `ExportToPackage` directly against the template — this script first creates a named DMF project (`DataManagementDefinitionGroups`) populated with one entity record per template line, then exports that project.

Use this script when you need a persistent, inspectable DMF project in D365 that matches the template structure, or when `Invoke-TemplateExport.ps1` is not available for the template you want.

**Per-template flow:**
1. Fetch all lines from `DefinitionGroupTemplateLines` for the selected template.
2. Delete any existing DMF project named `"<TemplateId> <LegalEntityId>"` (404 is silently ignored).
3. Create a fresh export project with that name.
4. `POST` one entity record to `DataManagementDefinitionGroupDetails` per template line.
5. Submit `ExportToPackage`, poll for completion, download the zip.

#### Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-EnvironmentUrl` | Yes | — | D365 base URL |
| `-TenantId` | Yes | — | Entra tenant ID or domain |
| `-LegalEntityId` | No | prompted | D365 company to export from; prompted interactively if omitted |
| `-TemplateName` | No | — | Process exactly this template (TemplateId) without showing the menu |
| `-DownloadPath` | No | `$env:TEMP` | Directory to save downloaded zip files; pass `''` to get the URL only |
| `-LogPath` | No | auto | Transcript log path; pass `''` to suppress |
| `-PollIntervalSeconds` | No | `30` | Status check interval (5–300) |
| `-TimeoutMinutes` | No | `60` | Per-template polling timeout (1–480) |
| `-MaxRetries` | No | `3` | Retry limit for transient REST failures (0–10) |
| `-Force` | No | off | Skip the confirmation prompt |
| `-WhatIf` | No | off | With `-TemplateName`: no API calls at all. Without: fetches the template list (read-only) for the menu, then exits without creating or exporting |
| `-PassThru` | No | off | Emit result objects: `Template, TemplateId, ProjectName, LegalEntityId, Status, LinesAdded, ExecutionId, DownloadUrl, DownloadedTo, Elapsed` |

#### Examples

```powershell
# Interactive: list templates, prompt for selection and legal entity
.\Invoke-ProjectExport.ps1 `
    -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com'
```

```powershell
# Supply legal entity up front, select templates interactively
.\Invoke-ProjectExport.ps1 `
    -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'USMF'
```

```powershell
# Non-interactive: one template, skip confirmation, download to C:\DMF\Downloads
.\Invoke-ProjectExport.ps1 `
    -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'USMF' `
    -TemplateName   '010 - System Setup' `
    -DownloadPath   'C:\DMF\Downloads' `
    -Force
```

```powershell
# WhatIf with a named template — zero API calls
.\Invoke-ProjectExport.ps1 `
    -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'USMF' `
    -TemplateName   '010 - System Setup' `
    -WhatIf
```

```powershell
# Export multiple templates and capture results
$results = .\Invoke-ProjectExport.ps1 `
    -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'USMF' `
    -DownloadPath   'C:\DMF\Downloads' `
    -Force `
    -PassThru

$results | Select-Object ProjectName, LinesAdded, Status, DownloadedTo
```

---

### Expand-ExportedPackages.ps1

Extracts downloaded DMF package zip files into individual subfolders.  Each subfolder becomes a self-contained package directory (Manifest.xml + xlsx files) that `Invoke-BaselineImport.ps1` can import directly.

This is the **review step** in the migration pipeline: after extraction you can open, edit, or filter the xlsx files before re-importing into a target environment.

No D365 API calls are made — this script is entirely local.

#### Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-SourcePath` | Yes | — | Directory containing downloaded `.zip` files |
| `-DestinationPath` | No | `SourcePath` | Root directory where per-package subfolders are created |
| `-PackageName` | No | — | Extract exactly this zip (with or without `.zip` extension) without showing the menu |
| `-Force` | No | off | Delete and re-extract if the destination folder already exists |
| `-WhatIf` | No | off | Show what would be extracted without writing any files |
| `-PassThru` | No | off | Emit result objects: `Package, DefinitionGroupId, ExtractedTo, XlsxCount, EntityCount, Status` |

Destination subfolders are named after the zip filename (without extension), preserving the timestamp so multiple exports of the same template do not collide.  Existing folders are shown in the menu in cyan; they are skipped unless `-Force` is specified.

#### Examples

```powershell
# Interactive: list all zips in the download folder and choose which to extract
.\Expand-ExportedPackages.ps1 `
    -SourcePath      'C:\DMF\Downloads' `
    -DestinationPath 'C:\DMF\Packages'
```

```powershell
# Extract everything, overwriting any folders that already exist
.\Expand-ExportedPackages.ps1 `
    -SourcePath      'C:\DMF\Downloads' `
    -DestinationPath 'C:\DMF\Packages' `
    -Force
```

```powershell
# Extract a single zip
.\Expand-ExportedPackages.ps1 `
    -SourcePath      'C:\DMF\Downloads' `
    -DestinationPath 'C:\DMF\Packages' `
    -PackageName     'SystemSetupExport_20240101120000'
```

```powershell
# WhatIf: preview extraction targets without writing anything
.\Expand-ExportedPackages.ps1 `
    -SourcePath 'C:\DMF\Downloads' `
    -WhatIf
```

---

### Invoke-PackageUpload.ps1

Uploads pre-built DMF package zip files directly to D365 without rebuilding the manifest.  Use this when you have complete, ready-to-import zips — for example, packages downloaded directly from another environment or obtained from a third party.

The definition group ID used in the `ImportFromPackage` call is read from `Manifest.xml` inside each zip, so the package is imported exactly as assembled.

#### Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-EnvironmentUrl` | Yes | — | D365 base URL |
| `-TenantId` | Yes | — | Entra tenant ID or domain |
| `-LegalEntityId` | Yes | — | D365 company to import into |
| `-UploadPath` | Yes | — | Directory containing the `.zip` files to upload |
| `-PackageName` | No | — | Upload exactly this zip (with or without `.zip` extension) without showing the menu |
| `-LogPath` | No | auto | Transcript log path; pass `''` to suppress |
| `-PollIntervalSeconds` | No | `30` | Status check interval (5–300) |
| `-TimeoutMinutes` | No | `60` | Per-package polling timeout (1–480) |
| `-MaxRetries` | No | `3` | Retry limit for transient REST failures (0–10) |
| `-Force` | No | off | Skip the confirmation prompt |
| `-WhatIf` | No | off | Validate packages locally without making any API calls |
| `-NoOverwrite` | No | off | Preserve existing D365 records |
| `-PassThru` | No | off | Emit result objects: `Package, DefinitionGroupId, Status, ExecutionId, Elapsed` |

#### Examples

```powershell
# Interactive: list all zips in the folder and choose which to upload
.\Invoke-PackageUpload.ps1 `
    -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'DAT' `
    -UploadPath     'C:\DMF\Downloads'
```

```powershell
# Upload one specific package, skip confirmation, preserve existing records
.\Invoke-PackageUpload.ps1 `
    -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'USMF' `
    -UploadPath     'C:\DMF\Downloads' `
    -PackageName    'SystemSetupExport_20240101120000.zip' `
    -NoOverwrite `
    -Force
```

```powershell
# Dry-run: validate all zips in the directory without connecting to D365
.\Invoke-PackageUpload.ps1 `
    -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'DAT' `
    -UploadPath     'C:\DMF\Downloads' `
    -WhatIf
```

---

## Full migration pipeline

Use this four-step workflow to copy data from one D365 environment (or legal entity) to another, with a local review stage in the middle.

```
Source environment / legal entity
        │
        │  Step 1 — Export
        ▼
C:\DMF\Downloads\  (zip files)
        │
        │  Step 2 — Extract
        ▼
C:\DMF\Packages\   (subfolders with Manifest.xml + xlsx)
        │
        │  Step 3 — Review / edit xlsx files
        ▼
C:\DMF\Packages\   (reviewed)
        │
        │  Step 4 — Import
        ▼
Target environment / legal entity
```

### Step 1 — Export from source

**Option A — export directly from a template** (simpler; uses the template as the `definitionGroupId`):

```powershell
.\Invoke-TemplateExport.ps1 `
    -EnvironmentUrl 'https://source.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'DAT' `
    -DownloadPath   'C:\DMF\Downloads' `
    -Force
```

**Option B — build a project from template lines, then export** (creates a named DMF project in D365 before exporting):

```powershell
.\Invoke-ProjectExport.ps1 `
    -EnvironmentUrl 'https://source.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'DAT' `
    -DownloadPath   'C:\DMF\Downloads' `
    -Force
```

All templates are listed in an interactive menu.  Select the ones you want (individual numbers, ranges like `1-5`, comma-separated list, or `A` for all).  Each selected template is exported and downloaded as a zip.

### Step 2 — Extract for review

```powershell
.\Expand-ExportedPackages.ps1 `
    -SourcePath      'C:\DMF\Downloads' `
    -DestinationPath 'C:\DMF\Packages' `
    -Force
```

Each zip becomes a subfolder under `C:\DMF\Packages`.  The subfolder contains `Manifest.xml`, `PackageHeader.xml`, and one xlsx file per entity.

### Step 3 — Review and edit

Open the xlsx files in any spreadsheet application and make changes as needed — filter rows, update values, add or remove records.  The manifest does not need to be edited; `Invoke-BaselineImport.ps1` rebuilds it automatically.

### Step 4 — Import into target

```powershell
.\Invoke-BaselineImport.ps1 `
    -EnvironmentUrl 'https://target.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'USMF' `
    -ResourcesPath  'C:\DMF\Packages'
```

The script scans `C:\DMF\Packages`, presents a selection menu, builds upload zips with optimised entity ordering, and imports each one into the target legal entity.

---

### Direct zip transfer (skip the review step)

If you want to copy packages from one environment to another without opening the files, use `Invoke-PackageUpload.ps1` directly on the downloaded zips:

```powershell
# Step 1 — export and download  (same as above)
.\Invoke-TemplateExport.ps1 `
    -EnvironmentUrl 'https://source.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'DAT' `
    -DownloadPath   'C:\DMF\Downloads' `
    -Force

# Step 2 — upload directly (no extraction needed)
.\Invoke-PackageUpload.ps1 `
    -EnvironmentUrl 'https://target.operations.dynamics.com' `
    -TenantId       'contoso.onmicrosoft.com' `
    -LegalEntityId  'USMF' `
    -UploadPath     'C:\DMF\Downloads' `
    -Force
```

---

## Entity execution ordering

`Invoke-BaselineImport.ps1` applies an execution ordering to each entity in the manifest before uploading.  This ensures that prerequisite entities (currencies, legal entities, address formats) are imported before the records that depend on them.

Each entity is assigned three values:

| Value | Meaning |
|---|---|
| `ExecutionUnit` (EU) | Processing group — all levels in one EU complete before the next EU starts |
| `LevelInExecutionUnit` (LV) | Sequential level within an EU |
| `SequenceInLevel` (SEQ) | Parallel execution slot within a level — entities sharing the same EU+LV+SEQ run concurrently |

The built-in defaults cover the most common D365 foundation entities.  Any entity not found in the defaults table keeps the values from its `Manifest.xml`.

### Customising order with ordering.json

Place an `ordering.json` file inside a package folder to override or extend the built-in defaults for that specific package.  Package-level entries take precedence over built-in values.

```json
{
  "My Custom Entity": { "EU": 1, "LV": 20, "SEQ": 60 },
  "Another Entity":   { "EU": 1, "LV": 20, "SEQ": 70 }
}
```

The file is optional.  If absent, only built-in defaults are applied.

---

## Library modules

The `lib/` folder contains helpers that are dot-sourced by the main scripts.  They are not intended to be called directly.

### DmfOutput.ps1

Console output and transcript helpers.

| Function | Description |
|---|---|
| `Write-Banner` | Prints the script title banner |
| `Write-Rule` | Prints a horizontal rule with an optional label |
| `Write-Step` | Cyan section heading |
| `Write-Info` | Gray informational line |
| `Write-Detail` | Dark gray detail line |
| `Write-OK` | Green success message |
| `Write-Warn` | Yellow warning message |
| `Write-Fail` | Red failure message |
| `Format-Elapsed` | Formats a `TimeSpan` as `2h 5m 30s` / `12m 4s` / `45s` |
| `Stop-RunTranscript` | Stops the PowerShell transcript if one is active |

### DmfRequest.ps1

`Invoke-DmfRequest` — wraps `Invoke-RestMethod` with:
- Automatic retry with exponential back-off for HTTP 5xx, 408, 429, and network errors
- `Retry-After` header support
- OData / D365 error detail extraction from response bodies
- Non-retryable treatment of HTTP 401 and other 4xx responses

### DmfPackage.ps1

| Export | Description |
|---|---|
| `$entityOrdering` | Ordered hashtable of built-in EU/LV/SEQ defaults |
| `Get-PackageInfo` | Inspects a package folder and returns metadata |
| `Resolve-EntityOrdering` | Merges built-in defaults with a package's `ordering.json` |

### DmfZip.ps1

`Get-ZipPackageInfo` — opens a DMF zip and returns metadata read from the embedded `Manifest.xml`: definition group name, entity count, and validation status.  Used by both `Invoke-PackageUpload.ps1` and `Expand-ExportedPackages.ps1`.
