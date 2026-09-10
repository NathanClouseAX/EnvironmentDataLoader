#Requires -Version 5.1
<#
.SYNOPSIS
    Generates "<template> (OData)" variants holding only the entities that can
    actually be read over OData.

.DESCRIPTION
    Roughly a third of the entities in the shipped templates are not
    OData-enabled, so -Mode OData skips them and reports NotPublic.  This
    script writes a companion template per source template containing only the
    entities that do work, so an OData run is clean and the DMF path is used
    knowingly for the rest.

    Per source template line:

      - the entity is OData-enabled  -> kept, with TargetEntity filled in from
        the catalog so it resolves without a Metadata-service call
      - it is not, but resources/entity-alternates.json confirms a substitute
        -> the substitute is written in its place and recorded
      - otherwise -> dropped, with the reason recorded

    Each generated folder gets a template.json marked "origin": "custom" (so
    Export-TemplateDefinition.ps1 never overwrites it) listing the counts,
    every substitution with its evidence, and every dropped entity with the
    reason.  A source template with no OData-capable entity generates nothing.

    Reads resources/entity-catalog.json, which is committed to this repository.
    Nothing here contacts D365 or Entra, and no development box is needed.

.PARAMETER ResourcesPath
    Folder holding the template folders, entity-catalog.json and
    entity-alternates.json.  Defaults to the repo-relative 'resources'.

.PARAMETER CatalogPath
    Explicit path to entity-catalog.json.  Defaults to ResourcesPath.

.PARAMETER AlternatesPath
    Explicit path to entity-alternates.json.  Optional; without it no
    substitutions are made.

.PARAMETER TemplateName
    Generate for this source template only.

.PARAMETER Suffix
    Folder-name suffix for the generated template.  Default ' (OData)'.

.PARAMETER Force
    Overwrite an existing generated folder.  Without it, existing folders are
    reported as Skipped-Exists.

.PARAMETER LogPath
    Transcript path.  Auto-generated in $env:TEMP when omitted; '' suppresses.

.PARAMETER WhatIf
    Report what would be generated without writing anything.

.PARAMETER PassThru
    Emit one object per source template: Template, Lines, Kept, Substituted,
    Dropped, Coverage, Folder, Status.

.EXAMPLE
    # Preview coverage for every template
    .\New-ODataTemplates.ps1 -WhatIf

.EXAMPLE
    # Generate them all, replacing any previous run
    .\New-ODataTemplates.ps1 -Force

.EXAMPLE
    .\New-ODataTemplates.ps1 -TemplateName '010 - System Setup' -Force
#>
[CmdletBinding()]
param(
    [string]$ResourcesPath,
    [string]$CatalogPath,
    [string]$AlternatesPath,
    [string]$TemplateName,
    [string]$Suffix = ' (OData)',

    [switch]$Force,
    [string]$LogPath,
    [switch]$WhatIf,
    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# =============================================================================
#  Load library modules
# =============================================================================
$libPath = Join-Path $PSScriptRoot 'lib'
. (Join-Path $libPath 'DmfOutput.ps1')
. (Join-Path $libPath 'DmfTemplate.ps1')
. (Join-Path $libPath 'DmfEntityCatalog.ps1')

$Script:Version   = '1.0'
$Script:LineWidth = 120

trap { if (Get-Command Stop-RunTranscript -ErrorAction SilentlyContinue) { Stop-RunTranscript }; break }

# =============================================================================
#  1.  Paths and inputs
# =============================================================================
if (-not $ResourcesPath)  { $ResourcesPath  = Join-Path $PSScriptRoot 'resources' }
if (-not $CatalogPath)    { $CatalogPath    = Join-Path $ResourcesPath 'entity-catalog.json' }
if (-not $AlternatesPath) { $AlternatesPath = Join-Path $ResourcesPath 'entity-alternates.json' }

if (-not (Test-Path -LiteralPath $ResourcesPath -PathType Container)) {
    throw "Resources path not found: '$ResourcesPath'"
}

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
if (-not $PSBoundParameters.ContainsKey('LogPath')) { $LogPath = Join-Path $env:TEMP "DMFODataTemplates_$stamp.log" }
if ($LogPath) {
    try { Start-Transcript -Path $LogPath -Force | Out-Null; $Script:TranscriptActive = $true } catch {}
}

Write-Banner -DryRun:$WhatIf -Title 'D365 F&O OData Template Generator'
Write-Info "Resources : $ResourcesPath"
Write-Info "Catalog   : $CatalogPath"
if ($WhatIf) { Write-Warn 'WhatIf active -- nothing will be written.' }

$catalogDocument = Get-DmfEntityCatalog -Path $CatalogPath
if ($null -eq $catalogDocument) {
    throw "Entity catalog not found at '$CatalogPath'. It ships with this repository under resources/; restore it or point -CatalogPath at a copy."
}

$catalog = @(ConvertTo-DmfCatalogEntities -Document $catalogDocument)
$sourceVersion = [string]$catalogDocument.sourceVersion
Write-Info "Catalog   : $($catalog.Count) entities, source version $(if ($sourceVersion) { $sourceVersion } else { 'unknown' })"
$lookup = New-DmfCatalogLabelLookup -Entities $catalog

# -- confirmed substitutions ---------------------------------------------------
$alternates = @{}
if (Test-Path -LiteralPath $AlternatesPath -PathType Leaf) {
    try {
        $altDocument = Get-Content -LiteralPath $AlternatesPath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($property in $altDocument.alternates.PSObject.Properties) {
            if ([string]$property.Value.status -ne 'confirmed') { continue }
            $alternates[$property.Name] = $property.Value
        }
        Write-Info "Alternates: $($alternates.Count) confirmed substitution(s)"
    } catch { Write-Warn "entity-alternates.json could not be parsed: $($_.Exception.Message)" }
} else {
    Write-Info 'Alternates: none (entity-alternates.json not present)'
}

# =============================================================================
#  2.  Generate
# =============================================================================
$scriptStart = Get-Date
$results = [System.Collections.Generic.List[object]]::new()

$folders = @(Get-TemplateFolders -ResourcesPath $ResourcesPath |
    Where-Object { $_.Name -notlike "*$Suffix" })
if ($TemplateName) {
    $folders = @($folders | Where-Object { $_.Name -eq $TemplateName })
    if ($folders.Count -eq 0) { throw "Template '$TemplateName' not found under '$ResourcesPath'." }
}
Write-Info "Templates : $($folders.Count) source template(s)"

foreach ($folder in $folders) {
    $manifestPath = Join-Path $folder.FullName 'Manifest.xml'
    try { $manifest = Read-DmfManifest -Path $manifestPath }
    catch { Write-Warn "$($folder.Name): $($_.Exception.Message)"; continue }

    $sourceLines = @($manifest.Lines)
    $keep = [System.Collections.Generic.List[object]]::new()
    $subs = [System.Collections.Generic.List[object]]::new()
    $drop = [System.Collections.Generic.List[object]]::new()

    foreach ($line in $sourceLines) {
        $label  = [string]$line.EntityName
        $entity = Find-DmfCatalogEntity -Label $label -Index $lookup
        $useLabel = $label
        $useTarget = ''

        if ($null -ne $entity -and $entity.IsPublic) {
            $useTarget = $entity.Name
        }
        elseif ($alternates.ContainsKey($label)) {
            $alternate = $alternates[$label]
            $useLabel  = [string]$alternate.entity
            $useTarget = [string]$alternate.targetEntity
            $subs.Add([ordered]@{
                from         = $label
                to           = $useLabel
                targetEntity = $useTarget
                rootTable    = [string]$alternate.rootTable
                fieldJaccard = $alternate.fieldJaccard
            })
        }
        else {
            $reason = if ($null -eq $entity) { 'not found in the entity catalog (ISV or renamed)' }
                      else { "not OData-enabled (IsPublic=false, $($entity.Name))" }
            $drop.Add([ordered]@{ entity = $label; reason = $reason })
            continue
        }

        $keep.Add([pscustomobject]@{
            EntityName               = $useLabel
            TargetEntity             = $useTarget
            ExecutionUnit            = [int]$line.ExecutionUnit
            LevelInExecutionUnit     = [int]$line.LevelInExecutionUnit
            SequenceInLevel          = [int]$line.SequenceInLevel
            FailLevelOnError         = [bool]$line.FailLevelOnError
            FailExecutionUnitOnError = [bool]$line.FailExecutionUnitOnError
            RunBusinessLogic         = [bool]$line.RunBusinessLogic
            RunBusinessValidation    = [bool]$line.RunBusinessValidation
            Disable                  = [bool]$line.Disable
        })
    }

    $coverage = $(if ($sourceLines.Count -gt 0) { [math]::Round(100 * $keep.Count / $sourceLines.Count) } else { 0 })
    $newName  = "$($folder.Name)$Suffix"
    $target   = Join-Path $ResourcesPath $newName
    $status   = 'Generated'

    if ($keep.Count -eq 0) { $status = 'Skipped-NoCoverage' }
    elseif ((Test-Path -LiteralPath (Join-Path $target 'Manifest.xml') -PathType Leaf) -and -not $Force) { $status = 'Skipped-Exists' }
    elseif ($WhatIf) { $status = 'WhatIf' }
    else {
        if (-not (Test-Path -LiteralPath $target -PathType Container)) {
            New-Item -ItemType Directory -Path $target -Force | Out-Null
        }
        $description = "$($manifest.Description)".Trim()
        $description = $(if ($description) { "$description -- OData-capable subset" } else { 'OData-capable subset' })

        $document = New-DmfManifestDocument -DefinitionGroupName $newName -Description $description -Lines $keep.ToArray()
        Write-DmfManifest      -Document $document -Path (Join-Path $target 'Manifest.xml')
        Write-DmfPackageHeader -Path (Join-Path $target 'PackageHeader.xml') -Description $description

        Write-TemplateSidecar -Folder $target -Sidecar ([ordered]@{
            schemaVersion  = 1
            templateId     = $newName
            origin         = 'custom'
            appliesTo      = 'OData'
            description    = $description
            sourceTemplate = $folder.Name
            generatedBy    = "New-ODataTemplates.ps1 $($Script:Version) (entity catalog $(if ($sourceVersion) { $sourceVersion } else { 'unknown' }))"
            generatedAt    = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
            notes          = 'Generated. Holds only the entities reachable over OData, with TargetEntity from the entity catalog. Entities under "dropped" have no public entity and must use the DMF path. Re-run New-ODataTemplates.ps1 -Force after changing the source template.'
            counts         = [ordered]@{ source = $sourceLines.Count; kept = $keep.Count; substituted = $subs.Count; dropped = $drop.Count }
            substituted    = $subs.ToArray()
            dropped        = $drop.ToArray()
        })
    }

    $results.Add([pscustomobject]@{
        Template = $folder.Name; Lines = $sourceLines.Count; Kept = $keep.Count
        Substituted = $subs.Count; Dropped = $drop.Count; Coverage = "$coverage%"
        Folder = $(if ($status -eq 'Generated') { $target } else { '-' }); Status = $status
    })
}

# =============================================================================
#  3.  Summary
# =============================================================================
Write-Host ''
Write-Rule "Result  --  $($results.Count) template(s)  |  $(Format-Elapsed ((Get-Date) - $scriptStart))"
$format = '  {0,-52} {1,6} {2,6} {3,6} {4,8} {5,9}  {6}'
Write-Host ($format -f 'Template', 'Lines', 'Kept', 'Subst', 'Dropped', 'Coverage', 'Status') -ForegroundColor White
Write-Host ($format -f ('-' * 52), '------', '------', '------', '--------', '---------', '------') -ForegroundColor DarkGray
foreach ($row in ($results | Sort-Object { [int]($_.Coverage -replace '%', '') })) {
    $colour = switch ($row.Status) {
        'Generated'          { if ($row.Dropped -eq 0) { 'Green' } else { 'White' } }
        'Skipped-NoCoverage' { 'Yellow' }
        'Skipped-Exists'     { 'DarkGray' }
        default              { 'Gray' }
    }
    $name = if ($row.Template.Length -gt 52) { $row.Template.Substring(0, 49) + '...' } else { $row.Template }
    Write-Host ($format -f $name, $row.Lines, $row.Kept, $row.Substituted, $row.Dropped, $row.Coverage, $row.Status) -ForegroundColor $colour
}

$totalLines = ($results | Measure-Object Lines -Sum).Sum
$totalKept  = ($results | Measure-Object Kept -Sum).Sum
$totalDrop  = ($results | Measure-Object Dropped -Sum).Sum
$totalSubs  = ($results | Measure-Object Substituted -Sum).Sum
Write-Host ''
Write-Info "$totalLines source line(s) -> $totalKept kept, $totalSubs substituted, $totalDrop dropped"
Write-Info "generated: $(@($results | Where-Object { $_.Status -eq 'Generated' }).Count)   no OData coverage: $(@($results | Where-Object { $_.Status -eq 'Skipped-NoCoverage' }).Count)   already present: $(@($results | Where-Object { $_.Status -eq 'Skipped-Exists' }).Count)"
if (@($results | Where-Object { $_.Status -eq 'Skipped-Exists' }).Count -gt 0) { Write-Info 'Re-run with -Force to regenerate existing folders.' }
if ($WhatIf) { Write-Info 'No files written. Remove -WhatIf to generate.' }

if ($PassThru) { $results.ToArray() }
Stop-RunTranscript
