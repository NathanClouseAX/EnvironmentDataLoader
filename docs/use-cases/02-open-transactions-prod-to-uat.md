# Use case 2 — Copy open transactions from production into UAT

Testers need realistic, current work in a UAT environment: the sales orders that are still open, the purchase orders still to be received, and the transfer orders that have been created but not shipped.  This guide takes those from **prod** and imports them into **UAT** with the template `resources/900 - Open transactions`.

It assumes UAT already has the master data those orders reference (customers, vendors, products, sites and warehouses, financial dimensions).  A UAT that was refreshed from prod has it; a UAT built from setup packages needs the master data imported first.  The order-status filters are the heart of this scenario, so read [where the filters live](#where-the-filters-live) before running anything.

Every command below is run from the repository root.

---

## The template

`resources/900 - Open transactions/Manifest.xml` is hand-authored — its `template.json` says `"origin": "custom"`, so menus tag it and the capture script never overwrites it (see [Creating a custom package](../custom-packages.md)) — and lists six entities, headers before lines:

| Entity (DMF label) | AOT entity | "Open" means | Filter to apply in the export project |
|---|---|---|---|
| Sales order headers V2 | `SalesOrderHeaderV2Entity` | not fully invoiced, not cancelled | **Sales order status** = *Open order* (optionally **Order type** = *Sales order* to leave out journals and returns) |
| Sales order lines V2 | `SalesOrderLineV2Entity` | line still has quantity to deliver or invoice | **Sales order line status** = *Open order* |
| Purchase order headers V2 | `PurchPurchaseOrderHeaderV2Entity` | not fully invoiced, not cancelled | **Purchase order status** = *Open order* |
| Purchase order lines V2 | `PurchPurchaseOrderLineV2Entity` | line still has quantity to receive or invoice | **Purchase order line status** = *Open order* |
| Transfer order headers | `InventTransferOrderHeaderEntity` | created, nothing shipped yet | **Transfer order status** = *Created* |
| Transfer order lines | `InventTransferOrderLineEntity` | line not shipped | **Transfer status** = *Created* |

Field labels can differ slightly between versions; the filter dialog in D365 shows the exact names.  Charges, notes and delivery schedules have their own entities (search *Data management → Data entities* for "charges") and can be added as further lines at level 20 if the tests need them.

Validate the template locally — no sign-in, no API calls:

```powershell
.\Invoke-ProjectExport.ps1 -EnvironmentUrl 'https://contoso-prod.operations.dynamics.com' -TenantId 'contoso.onmicrosoft.com' -LegalEntityId 'USMF' -TemplateName '900 - Open transactions' -WhatIf
```

---

## Where the filters live

A DMF template holds **entities and their order, not filters**.  Filters belong to the *data project* — the *Filter* button on each entity line in *Data management → Export*.  `Invoke-ProjectExport.ps1` builds its project from the template lines and exports immediately, so a project it creates has **no filters**: run it against a production company with years of history and it exports every sales order ever entered.

Two ways to get a filtered export, pick one:

**Path A — build the filtered project once in prod, export it repeatedly (recommended).**

1. In prod: *Data management → Export*, create a project named, say, `Open transactions USMF`, add the six entities in the order above (headers at level 10, lines at level 20), and set the status filters from the table on each line.  Save it.  This is a one-time task; the project stays in prod.
2. Export it with `Invoke-TemplateExport.ps1`.  Its `-TemplateName` accepts a **data project name**, and `ExportToPackage` honours the project's filters:

```powershell
$tenant = 'contoso.onmicrosoft.com'
$prod   = 'https://contoso-prod.operations.dynamics.com'
$uat    = 'https://contoso-uat.sandbox.operations.dynamics.com'

.\Invoke-TemplateExport.ps1 -EnvironmentUrl $prod -TenantId $tenant -LegalEntityId 'USMF' `
    -TemplateName 'Open transactions USMF' -DownloadPath 'C:\DMF\open-orders' -PollIntervalSeconds 10 -Force
```

The zip lands in `C:\DMF\open-orders`.  Do **not** point `Invoke-ProjectExport.ps1` at a project of the same name: it deletes and recreates any project that matches `<template> <company>`, and the filters go with it.

**Path B — export everything, trim locally (small companies only).**

If the company is small enough that a full export is acceptable, let the template do the work and filter the xlsx afterwards:

```powershell
.\Invoke-ProjectExport.ps1 -EnvironmentUrl $prod -TenantId $tenant -LegalEntityId 'USMF' `
    -TemplateName '900 - Open transactions' -DownloadPath 'C:\DMF\open-orders' -PollIntervalSeconds 10 -Force
```

Then, in `C:\DMF\open-orders\900-Open-transactions-USMF_<timestamp>\`, delete every row whose status column is not *Open order* / *Created* in each of the six files (a spreadsheet filter, or `Import-Excel` / `Export-Excel` from the ImportExcel module).  Keep header and line files consistent: a line whose header you removed will fail in staging.

---

## Step 1 — Review the package

For Path A extract the zip first; Path B's folder is already extracted.

```powershell
.\Expand-ExportedPackages.ps1 -SourcePath 'C:\DMF\open-orders' -DestinationPath 'C:\DMF\open-orders\packages' -Force
```

Things to look at in the xlsx files before importing:

- **Partial deliveries.**  The import creates orders with the *ordered* quantity; delivered and invoiced quantities are history that DMF does not recreate.  If a test needs an order that is half-delivered, set the ordered quantity on that line to the remaining quantity.
- **Order numbers.**  Orders keep their production numbers.  In a UAT refreshed from prod the same numbers already exist and the import updates them (default overwrite); add `-NoOverwrite` to leave existing orders alone and only insert the ones UAT does not have.  In a UAT built from scratch, make sure the sales and purchase order number sequences allow manual numbers, or the import rejects the supplied ones.
- **Dates.**  Requested ship and receipt dates are copied as they are; move them forward if the tests depend on "future" orders.
- **Personal data.**  Sales orders carry customer names and addresses.  Apply whatever masking your UAT policy requires before importing.
- **Purchase order state.**  Imported purchase orders arrive as new orders.  Whether they are confirmed or waiting for approval depends on UAT's change-management settings; plan on confirming them there.

---

## Step 2 — Dry run, then import into UAT

Check the target URL twice: the import scripts do exactly what they are told, and UAT is the only place this package should go.

```powershell
# Plan only
.\Invoke-BaselineImport.ps1 -EnvironmentUrl $uat -TenantId $tenant -LegalEntityId 'USMF' -ResourcesPath 'C:\DMF\open-orders\packages' -WhatIf

# Import (headers before lines is preserved from the manifest)
.\Invoke-BaselineImport.ps1 -EnvironmentUrl $uat -TenantId $tenant -LegalEntityId 'USMF' -ResourcesPath 'C:\DMF\open-orders\packages' -PollIntervalSeconds 10 -Force
```

`Invoke-BaselineImport.ps1` rebuilds the manifest with its built-in ordering for the foundation entities it knows; the order entities are not in that table, so the level-10 / level-20 split from the export is kept.

If nothing needs editing, the zip from Path A can go straight in without extraction:

```powershell
.\Invoke-PackageUpload.ps1 -EnvironmentUrl $uat -TenantId $tenant -LegalEntityId 'USMF' -UploadPath 'C:\DMF\open-orders' -Force
```

---

## Step 3 — Check the results

```powershell
.\Get-ExecutionJobReport.ps1 -EnvironmentUrl $uat -TenantId $tenant -IssuesOnly -HtmlPath 'C:\DMF\open-orders\import-issues.html'
```

Typical failures and their fixes:

| Symptom | Cause | Fix |
|---|---|---|
| Header imports, lines fail with "order does not exist" | header rejected earlier, or lines ran before headers | fix the header error first; check the manifest keeps lines at a later level |
| "Item/customer/vendor/warehouse does not exist" | master data missing in UAT | import the master data first (customers, vendors, released products, warehouses), then re-run |
| "Number sequence … does not allow manual" | order number sequence is continuous / not manual in UAT | allow manual numbers on the sequence, or blank the order-number column to let UAT assign new ones (then lines must reference the new numbers — use overwrite on a refreshed UAT instead) |
| Dimension validation errors | dimension values or account structures differ | align the setup first ([use case 1](01-master-data-prod-to-prod.md)) |

Re-run Step 2 for the affected package after fixing the xlsx.

---

## Verifying what arrived

The DMF job report is the authoritative check.  For a quick count, an OData pull of the same six entities works but has no status filter either, so cap it and treat the numbers as a spot check rather than a reconciliation:

```powershell
.\Invoke-ProjectExport.ps1 -EnvironmentUrl $uat -TenantId $tenant -LegalEntityId 'USMF' -TemplateName '900 - Open transactions' -Mode OData -MaxRecordsPerEntity 5000 -DataPath 'C:\DMF\snapshots' -Force
```

---

## Refreshing UAT again later

Path A makes a refresh a two-command job: export the saved project from prod, import (or upload) into UAT.  The filters stay in prod's project; the template in `resources/` documents the entity set and keeps the `-WhatIf` plan and the OData spot check available.  If you refresh often, add `-NoOverwrite` on the import to leave orders that testers have already worked on untouched.
