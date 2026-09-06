# Creating a custom package

A **custom package** is a folder under `resources/` that you author and maintain yourself rather than capture from an environment.  It starts life as a **template** (a `Manifest.xml` listing entities and their execution order) and becomes a **package** the moment `.xlsx` data files sit next to the manifest.  The same folder drives every script: the DMF export, the OData pull, the compare, and the import.

`resources/900 - Open transactions` is the worked example; [use case 2](use-cases/02-open-transactions-prod-to-uat.md) shows it in action.  Every command below runs from the repository root and, unless it says otherwise, makes no API calls.

---

## 1. Decide what goes in

**Entities.**  Open *Data management → Data entities* in D365.  The manifest uses the **entity label** shown in the *Entity name* column (for example `Sales order headers V2`); the *Target entity* column is the AOT name (`SalesOrderHeaderV2Entity`) and is optional in the manifest, but worth adding because it lets the OData path skip the label lookup.  Entity labels are language-specific; write them in the language your environment runs the Data management workspace in.

**Order.**  DMF runs a package by execution unit, then level, then sequence:

| Value | Meaning |
|---|---|
| `ExecutionUnit` | Group.  All levels of unit 1 finish before unit 2 starts. |
| `LevelInExecutionUnit` | Step inside a unit.  Level 10 finishes before level 20 starts. |
| `SequenceInLevel` | Entities with the same unit, level **and** sequence run in parallel; different sequences run one after another. |

Parents before children: headers at level 10, lines at level 20; groups before the records that reference them.  When in doubt, copy the values from the Microsoft template that ships the same entity (`resources/<template>/Manifest.xml` — those files are UTF-16, so open them in an editor that shows UTF-16 correctly, or read them with PowerShell's `[xml](Get-Content -Raw …)`).

**Scope.**  A template lists entities, not filters.  Row filters belong to the export *data project* in D365 (see [section 6](#6-filters-and-scope)).

---

## 2. Create the folder and the manifest

The folder name is the **template ID**: it appears in menus, is what `-TemplateName` matches, becomes the DMF project name (`<template> <company>`), and prefixes the export folder.  Conventions that work well:

- number custom templates from `900` upwards so they sort after Microsoft's `010 … 650`;
- avoid characters that are invalid in file names (`\ / : * ? " < > |`) — D365 drops them from file names, and so do the scripts (`Country/regions` becomes `Countryregions`);
- keep the name stable once it has been used, because export folders and DMF projects carry it.

Write `resources/<template ID>/Manifest.xml`:

```xml
<?xml version="1.0" encoding="utf-8"?>
<DataManagementPackageManifest xmlns:i="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/dynamics/2015/01/DataManagement">
  <DefinitionGroupName>900 - Open transactions</DefinitionGroupName>
  <Description>Open sales orders, open purchase orders and transfer orders not yet shipped</Description>
  <PackageEntityList>
    <DataManagementPackageEntityData>
      <EntityName>Sales order headers V2</EntityName>
      <TargetEntity>SalesOrderHeaderV2Entity</TargetEntity>
      <ExecutionUnit>1</ExecutionUnit>
      <LevelInExecutionUnit>10</LevelInExecutionUnit>
      <SequenceInLevel>10</SequenceInLevel>
    </DataManagementPackageEntityData>
    <DataManagementPackageEntityData>
      <EntityName>Sales order lines V2</EntityName>
      <TargetEntity>SalesOrderLineV2Entity</TargetEntity>
      <ExecutionUnit>1</ExecutionUnit>
      <LevelInExecutionUnit>20</LevelInExecutionUnit>
      <SequenceInLevel>10</SequenceInLevel>
    </DataManagementPackageEntityData>
    <DataManagementPackageEntityData>
      <EntityName>Sales order header charges</EntityName>
      <Disable>true</Disable>
      <ExecutionUnit>1</ExecutionUnit>
      <LevelInExecutionUnit>20</LevelInExecutionUnit>
      <SequenceInLevel>20</SequenceInLevel>
    </DataManagementPackageEntityData>
  </PackageEntityList>
</DataManagementPackageManifest>
```

Per line, only `EntityName` and the three ordering values are required.  Everything else is optional and defaults as follows:

| Element | Default | Notes |
|---|---|---|
| `TargetEntity` | none | AOT entity name; filled in by the entity resolver when the OData path needs it |
| `Disable` | `false` | `true` keeps the line in the file but out of exports and pulls (`-IncludeDisabled` restores it) |
| `FailLevelOnError`, `FailExecutionUnitOnError` | `false` | Stop the level / the unit when this entity fails |
| `RunBusinessLogic`, `RunBusinessValidation` | `true` | Passed to the DMF project |
| `SourceFormat` | `EXCEL` | |
| `InputFilePath` | `<EntityName>.xlsx` | Invalid characters removed |
| `ExcelSheetName` | `<Entity_Name>$` | Spaces to underscores, trailing `$` |

Element order does not matter; the root namespace does.  UTF-8 or UTF-16 are both read; D365 writes UTF-16 LE with a BOM, and the scripts do the same when they regenerate a manifest.

---

## 3. Mark it as custom

Add `template.json` next to the manifest:

```json
{
  "schemaVersion": 1,
  "templateId": "900 - Open transactions",
  "origin": "custom",
  "description": "Open sales orders, open purchase orders and transfer orders not yet shipped",
  "notes": "Hand-authored. Headers at level 10, lines at level 20. Status filters must be set on the export data project."
}
```

`"origin": "custom"` is the field that matters; the rest is for the reader.  The marker does three things:

- **Menus show it.**  `Invoke-BaselineImport.ps1` lists the folder as `<name>  [custom]`; `Invoke-ProjectExport.ps1` shows `custom, N lines`.  `Get-TemplateInfo` and `Get-PackageInfo` expose `Origin` (`custom` / `captured` / `unknown`) and `IsCustom`.
- **Capture never overwrites it.**  `Export-TemplateDefinition.ps1` skips the folder with status `Skipped-Custom` even when `-Force` is given, so an environment template that happens to share the name cannot replace your work.
- **It documents intent.**  Folders the capture script writes carry `"origin": "captured"` plus `capturedFrom` / `capturedAt`; re-capture refreshes those.  A folder without a sidecar (an expanded export dropped under `resources/`) is `unknown`.

When *not* to mark: if you author the template in D365 itself and capture it, leave the sidecar as captured — D365 is the source of truth and re-capture should win.  Mark custom when the repository is the source of truth.

---

## 4. Validate without signing in

```powershell
$env = 'https://contoso-prod.operations.dynamics.com'; $tenant = 'contoso.onmicrosoft.com'

# structure, ordering, disabled lines, project name; shows "custom, N lines"
.\Invoke-ProjectExport.ps1 -EnvironmentUrl $env -TenantId $tenant -LegalEntityId 'USMF' -TemplateName '900 - Open transactions' -WhatIf

# how each entity would resolve for the OData path, from entity-map.json alone
.\Invoke-ProjectExport.ps1 -EnvironmentUrl $env -TenantId $tenant -LegalEntityId 'USMF' -TemplateName '900 - Open transactions' -Mode OData -WhatIf

# what the import would see (a template without xlsx files is reported as "Template only")
.\Invoke-BaselineImport.ps1 -EnvironmentUrl $env -TenantId $tenant -LegalEntityId 'USMF' -PackageName '900 - Open transactions' -WhatIf
```

Structural problems are reported as warnings with the folder name: wrong root element or namespace, no entity lines, a line without `EntityName`, a duplicate `EntityName`, a missing or non-integer ordering value.  A folder with warnings other than the namespace one is left out of the export menu until fixed.

An entity label that does not exist in the target environment is only detected online: the export logs *Failed to add entity '<label>'* for that line and carries on with the rest.

---

## 5. From template to package

The first real export turns the template into a package:

```powershell
.\Invoke-ProjectExport.ps1 -EnvironmentUrl $env -TenantId $tenant -LegalEntityId 'USMF' -TemplateName '900 - Open transactions' -DownloadPath 'C:\DMF\exports' -Force
```

This creates the DMF project `900 - Open transactions USMF` in the environment (deleting any project of that name first), exports it, downloads the zip, and extracts it to `C:\DMF\exports\900-Open-transactions-USMF_<timestamp>\`.  That folder is a complete package: a manifest with field maps, `PackageHeader.xml`, and one xlsx per entity.

Then:

1. **Seed the entity map** so the OData path and the compare know the AOT names and keys without asking the Metadata service:

   ```powershell
   .\Export-TemplateDefinition.ps1 -EnvironmentUrl $env -TenantId $tenant -SeedFromPath 'C:\DMF\exports'
   ```

2. **Make the repository folder a package** (optional): copy the xlsx files and `PackageHeader.xml` from the export into `resources/900 - Open transactions/`.  You may also replace your hand-written `Manifest.xml` with the exported one — it carries the field maps D365 generated — as long as `template.json` stays.  `Invoke-BaselineImport.ps1` then lists the folder as a full package instead of *Template only*.

3. **Import ordering** (optional): drop an `ordering.json` in the folder to override execution values at import time for specific entities; see *Entity execution ordering* in the README.  It affects `Invoke-BaselineImport.ps1` only.

Remember that `resources/` is committed and `data/` is not: a package with real business data in its xlsx files should stay out of git or be scrubbed first.

---

## 6. Filters and scope

A template cannot carry a row filter.  If the package must hold a subset (open orders, one site, one year), either

- build the export project once in D365 with the filters and export it with `Invoke-TemplateExport.ps1 -TemplateName '<project name>'`, which honours them — the pattern in [use case 2](use-cases/02-open-transactions-prod-to-uat.md); or
- export everything and trim the xlsx files locally before importing.

For OData pulls, `-MaxRecordsPerEntity` caps each entity and marks the file `Truncated`; there is no server-side filter either.

---

## Checklist

- [ ] Folder under `resources/`, named after the template ID, no invalid file-name characters
- [ ] `Manifest.xml` with the DataManagement namespace, one line per entity, headers before lines
- [ ] `TargetEntity` filled in where known
- [ ] `template.json` with `"origin": "custom"`
- [ ] The three `-WhatIf` runs in section 4 are clean
- [ ] `Invoke-Pester ./tests` still green (a template folder never changes tests, but a broken manifest shows up as warnings in every menu)
- [ ] README or the relevant use case mentions the template if others are expected to use it
