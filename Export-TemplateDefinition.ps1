#Requires -Version 5.1
<#
.SYNOPSIS
    Captures DMF templates from a D365 F&O environment into local template
    folders (resources/<TemplateId>/Manifest.xml), including Microsoft's
    default templates once they have been loaded in the environment.

.DESCRIPTION
    1. Optionally seeds resources/entity-map.json from any Manifest.xml files
       found under -SeedFromPath (no sign-in needed).  With no template
       selection this is the whole run.
    2. Authenticates via Microsoft Entra device code flow.
    3. Lists DefinitionGroupTemplateHeaders and presents a selection menu
       (skipped with -TemplateName or -All).
    4. Resolves every entity in the selected templates through the entity map
       and the Metadata service so each line gets its TargetEntity (AOT name)
       and the map gains the OData collection, key fields and company flag.
       Unresolved labels are reported but do not fail the capture (-NoResolve
       skips this step).
    5. For each template, reads DefinitionGroupTemplateLines and writes:
         resources/<TemplateId>/Manifest.xml       (UTF-16 LE, D365 format)
         resources/<TemplateId>/PackageHeader.xml
         resources/<TemplateId>/template.json      (provenance; -NoSidecar skips it)
       Existing folders are skipped unless -Force; .xlsx files are never touched.
       A folder whose template.json says "origin": "custom" is never
       overwritten, not even with -Force (status Skipped-Custom).
    6. Saves the entity map and prints a summary table.

    The captured folder is a template for Invoke-ProjectExport.ps1
    (-TemplateSource Local, both -Mode Dmf and -Mode OData) and, once .xlsx
    files are added, a package for Invoke-BaselineImport.ps1.

    Library files (in ./lib/)
    ─────────────────────────
    DmfOutput.ps1    -- Write-* helpers, Format-Elapsed, Stop-RunTranscript
    DmfRequest.ps1   -- Invoke-DmfRequest (REST client with retry)
    DmfAuth.ps1      -- Connect-DmfEnvironment, Get-DmfAuthHeaders, Test-DmfTokenExpiry
    DmfOData.ps1     -- Get-DmfODataAll, New-DmfODataUri
    DmfTemplate.ps1  -- New-DmfManifestDocument, Write-DmfManifest, sidecar helpers
    DmfMetadata.ps1  -- Resolve-DmfEntity, entity-map.json

.PARAMETER EnvironmentUrl
    D365 base URL, e.g. https://contoso.operations.dynamics.com

.PARAMETER TenantId
    Entra tenant ID or domain, e.g. contoso.onmicrosoft.com

.PARAMETER TemplateName
    Capture exactly this template (TemplateId); skips the menu.

.PARAMETER All
    Capture every validated template; skips the menu.  Mutually exclusive
    with -TemplateName.

.PARAMETER ResourcesPath
    Root under which <TemplateId>/ folders are written and where
    entity-map.json lives.  Default: ./resources

.PARAMETER SeedFromPath
    One or more folders to scan recursively for Manifest.xml files
    (e.g. expanded exports).  Every EntityName -> TargetEntity pair found is
    merged into entity-map.json before resolution.  Usable on its own.

.PARAMETER Force
    Overwrite the manifest, header and sidecar of an existing folder.  A
    folder marked "origin": "custom" in template.json is never overwritten
    (see docs/custom-packages.md).

.PARAMETER NoResolve
    Skip Metadata-service resolution; manifests carry the label and any
    TargetEntity already known from the map.

.PARAMETER NoSidecar
    Do not write template.json.

.PARAMETER Language
    Label language used for resolution.  Default: en-US

.PARAMETER AuthMode
    How to sign in.  Auto (default): open the default browser when the
    session is interactive, falling back to the device code flow if that
    fails; Browser: browser only; DeviceCode: print a code to enter in
    any browser (for SSH sessions and servers without a browser).

.PARAMETER LogPath
    Transcript path; auto-generated under $env:TEMP when omitted; '' suppresses.

.PARAMETER MaxRetries
    Retry limit for transient REST failures (0-10).  Default 3.

.PARAMETER WhatIf
    With -TemplateName: no API calls at all.  Otherwise the template list is
    fetched (read-only) so the menu can be shown; nothing is written.

.PARAMETER PassThru
    Emit one result object per template: TemplateId, Description, Lines,
    Resolved, Unresolved, Folder, Status, Elapsed.

.EXAMPLE
    # Capture every template (after Data management > Templates > Load default templates)
    .\Export-TemplateDefinition.ps1 -EnvironmentUrl 'https://contoso.operations.dynamics.com' -TenantId 'contoso.onmicrosoft.com' -All

.EXAMPLE
    # One template, overwrite an earlier capture
    .\Export-TemplateDefinition.ps1 -EnvironmentUrl 'https://contoso.operations.dynamics.com' -TenantId 'contoso.onmicrosoft.com' -TemplateName '010 - System Setup' -Force

.EXAMPLE
    # Seed the entity map from expanded exports without signing in
    .\Export-TemplateDefinition.ps1 -EnvironmentUrl 'https://contoso.operations.dynamics.com' -TenantId 'contoso.onmicrosoft.com' -SeedFromPath 'C:\DMF\Packages'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^https?://')]
    [string]$EnvironmentUrl,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [string]$TemplateName,
    [switch]$All,

    [string]$ResourcesPath,

    [string[]]$SeedFromPath,

    [switch]$Force,
    [switch]$NoResolve,
    [switch]$NoSidecar,

    [ValidateNotNullOrEmpty()]
    [string]$Language = 'en-US',

    [ValidateSet('Auto', 'Browser', 'DeviceCode')]
    [string]$AuthMode = 'Auto',

    [AllowEmptyString()]
    [string]$LogPath,

    [ValidateRange(0, 10)]
    [int]$MaxRetries = 3,

    [switch]$WhatIf,
    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# An uncaught error anywhere below must not leave the transcript running in
# the caller's console (it would silently swallow every later command's
# output into this log).  Stop it, then let the error propagate.
trap { if (Get-Command -Name Stop-RunTranscript -ErrorAction SilentlyContinue) { Stop-RunTranscript }; break }

# =============================================================================
#  Load library modules
# =============================================================================
$libPath = Join-Path $PSScriptRoot 'lib'
. (Join-Path $libPath 'DmfOutput.ps1')
. (Join-Path $libPath 'DmfRequest.ps1')
. (Join-Path $libPath 'DmfAuth.ps1')
. (Join-Path $libPath 'DmfOData.ps1')
. (Join-Path $libPath 'DmfTemplate.ps1')
. (Join-Path $libPath 'DmfMetadata.ps1')

# =============================================================================
#  Pre-flight validation
# =============================================================================
if ($All -and $TemplateName) {
    throw '-All and -TemplateName are mutually exclusive.'
}
# Resolved here, not in the param block: on PS 5.1 $PSScriptRoot is empty
# while defaults are evaluated for a script started with a relative path.
if (-not $ResourcesPath) { $ResourcesPath = Join-Path $PSScriptRoot 'resources' }
if (-not (Test-Path -LiteralPath $ResourcesPath -PathType Container)) {
    if ($WhatIf) { Write-Host "Resources path '$ResourcesPath' does not exist; it would be created." }
    else { New-Item -ItemType Directory -Path $ResourcesPath -Force | Out-Null }
}
# A seed-only run is: -SeedFromPath given and nothing selects templates.
$seedOnly = ($null -ne $SeedFromPath -and $SeedFromPath.Count -gt 0 -and -not $All -and -not $TemplateName)

# =============================================================================
#  Script-level constants
# =============================================================================
$Script:Version          = '1.0'
$Script:LineWidth        = 80
$Script:MaxRetries       = $MaxRetries
$Script:TranscriptActive = $false
try {
    $w = $Host.UI.RawUI.BufferSize.Width
    $Script:LineWidth = [Math]::Min(120, [Math]::Max(80, $w))
} catch { <# non-interactive host -- keep default #> }

# =============================================================================
#  1.  Transcript + banner
# =============================================================================
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
if (-not $PSBoundParameters.ContainsKey('LogPath')) { $LogPath = Join-Path $env:TEMP "DMFTemplateCapture_${ts}.log" }
if ($LogPath -ne '' -and -not $WhatIf) {
    try { Start-Transcript -Path $LogPath -Force | Out-Null; $Script:TranscriptActive = $true } catch {}
}

Write-Banner -DryRun:$WhatIf -Title 'D365 F&O Template Capture Utility'
Write-Host ''
Write-Info "Environment  : $EnvironmentUrl"
Write-Info "Tenant       : $TenantId"
Write-Info "Resources    : $ResourcesPath"
if ($SeedFromPath) { Write-Info "Seed from    : $($SeedFromPath -join ', ')" }
if ($Script:TranscriptActive) { Write-Info "Log          : $LogPath" }
if ($WhatIf) { Write-Warn 'WhatIf active -- nothing will be written.' }
if ($Force)  { Write-Info 'Force        : existing template folders will be overwritten' }

$mapPath = Join-Path $ResourcesPath 'entity-map.json'
$map     = Get-DmfEntityMap -Path $mapPath
Write-Info "Entity map   : $mapPath  ($($map.Entities.Count) entries)"

# =============================================================================
#  2.  Seed the entity map from manifests on disk  (no network)
# =============================================================================
if ($SeedFromPath) {
    Write-Step 'Harvesting EntityName -> TargetEntity pairs from manifests'
    $seed = Update-DmfEntityMapFromManifests -Path $SeedFromPath -Map $map
    Write-Info "$($seed.Files) manifest(s), $($seed.Lines) line(s): $($seed.Added) added, $($seed.Updated) updated, $($seed.AlreadyKnown) already known."
    if (-not $WhatIf -and ($seed.Added + $seed.Updated) -gt 0) {
        [void](Save-DmfEntityMap -Map $map)
        Write-Info "Entity map saved."
    }
    if ($seedOnly) {
        Write-OK 'Seed complete.  No templates selected, so no sign-in was needed.'
        Stop-RunTranscript
        return
    }
}

# =============================================================================
#  3.  WhatIf early exit for a named template  (zero API calls)
# =============================================================================
if ($WhatIf -and $TemplateName) {
    $folder = Join-Path $ResourcesPath (ConvertTo-DmfSafeFileName -Name $TemplateName)
    $state  = if (-not (Test-Path -LiteralPath $folder))                { '(new)' }
              elseif ((Get-TemplateOrigin -Folder $folder) -eq 'custom') { '(exists, marked custom in template.json -- never overwritten)' }
              elseif ($Force)                                           { '(exists, would overwrite)' }
              else                                                      { '(exists, would be skipped -- use -Force)' }
    Write-Host ''
    Write-Rule 'WhatIf -- nothing will be written'
    Write-Info "Would capture template : $TemplateName"
    Write-Info "Target folder          : $folder  $state"
    Write-Info 'No API calls made.  Remove -WhatIf to capture.'
    Stop-RunTranscript
    return
}

# =============================================================================
#  4.  Authenticate
# =============================================================================
Write-Step 'Authenticating with Microsoft Entra (device code flow)'
$baseUrl = $EnvironmentUrl.TrimEnd('/')
$session = Connect-DmfEnvironment -EnvironmentUrl $baseUrl -TenantId $TenantId -AuthMode $AuthMode
$Script:DmfSession = $session
$authHeaders       = Get-DmfAuthHeaders -Session $session

# =============================================================================
#  5.  List templates
# =============================================================================
Write-Step 'Fetching available templates'
$rawTemplates = @(Get-DmfODataAll -Uri "$baseUrl/data/DefinitionGroupTemplateHeaders" -Operation 'list templates' -Headers $authHeaders)
$allTemplates = @($rawTemplates |
    Where-Object { (Get-DmfProp $_ 'Status' '') -eq 'Validated' } |
    Sort-Object -Property TemplateId |
    ForEach-Object -Begin { $i = 1 } -Process {
        [pscustomobject]@{
            Index       = $i++
            TemplateId  = [string]$_.TemplateId
            Description = [string](Get-DmfProp $_ 'Description' '')
        }
    })
if ($allTemplates.Count -eq 0) { throw 'No validated templates found in this environment.' }
Write-Info "$($allTemplates.Count) validated template(s) found."

# =============================================================================
#  6.  Selection
# =============================================================================
$selected = [System.Collections.Generic.List[pscustomobject]]::new()
if ($TemplateName) {
    $match = $allTemplates | Where-Object { $_.TemplateId -eq $TemplateName }
    if (-not $match) {
        $available = ($allTemplates | ForEach-Object { "    '$($_.TemplateId)'" }) -join [System.Environment]::NewLine
        throw "Template '$TemplateName' not found.`nAvailable templates:`n$available"
    }
    $selected.Add($match)
}
elseif ($All) {
    foreach ($t in $allTemplates) { $selected.Add($t) }
}
else {
    Write-Host ''
    Write-Rule 'Available templates'
    foreach ($t in $allTemplates) {
        $exists = Test-Path -LiteralPath (Join-Path $ResourcesPath (ConvertTo-DmfSafeFileName -Name $t.TemplateId)) -PathType Container
        $tag    = if ($exists) { '  [local copy exists]' } else { '' }
        $desc   = if ($t.Description) { "  -- $($t.Description)" } else { '' }
        Write-Host ("  [{0,3}]  {1}{2}{3}" -f $t.Index, $t.TemplateId, $desc, $tag) -ForegroundColor $(if ($exists) { 'Cyan' } else { 'Gray' })
    }
    Write-Host ''
    Write-Info 'Enter numbers (e.g. 1,3,5-7), A for all, or Q to quit.'
    $indices = $null
    do {
        $rawInput = (Read-Host '  Selection').Trim().ToLower()
        if ($rawInput -in 'q', 'quit') { $indices = @(); break }
        if ($rawInput -in 'a', 'all', '') { $indices = 1..$allTemplates.Count; break }
        $parsed = [System.Collections.Generic.List[int]]::new(); $bad = $false
        foreach ($token in ($rawInput -split ',')) {
            $token = $token.Trim()
            if ($token -match '^(\d+)\s*-\s*(\d+)$') { ([int]$Matches[1])..([int]$Matches[2]) | ForEach-Object { $parsed.Add($_) } }
            elseif ($token -match '^\d+$') { $parsed.Add([int]$token) }
            else { Write-Warn "  Unrecognised input '$token'."; $bad = $true; break }
        }
        if ($bad) { continue }
        $out = @($parsed | Where-Object { $_ -lt 1 -or $_ -gt $allTemplates.Count })
        if ($out.Count -gt 0) { Write-Warn "  Out-of-range: $($out -join ', ')."; continue }
        if ($parsed.Count -eq 0) { Write-Warn '  Nothing selected.'; continue }
        $indices = @($parsed | Sort-Object -Unique); break
    } while ($true)
    if (-not $indices -or $indices.Count -eq 0) { Write-Warn 'No templates selected.  Exiting.'; Stop-RunTranscript; return }
    foreach ($idx in $indices) { $selected.Add(($allTemplates | Where-Object { $_.Index -eq $idx })) }
}
Write-Info "$($selected.Count) template(s) selected."

if ($WhatIf) {
    Write-Host ''
    Write-Rule 'WhatIf -- nothing will be written'
    foreach ($t in $selected) {
        $folder = Join-Path $ResourcesPath (ConvertTo-DmfSafeFileName -Name $t.TemplateId)
        Write-Detail "[$($t.Index)] $($t.TemplateId)  ->  $folder  $(if (Test-Path -LiteralPath $folder) { '(exists)' } else { '(new)' })"
    }
    Stop-RunTranscript
    return
}

# =============================================================================
#  7.  Fetch lines for every selected template, then resolve once
# =============================================================================
Write-Step 'Fetching template lines'
$linesByTemplate = @{}
foreach ($t in $selected) {
    $uri   = New-DmfODataUri -BaseUrl $baseUrl -Collection 'DefinitionGroupTemplateLines' -Filter "TemplateId eq '$(ConvertTo-DmfODataLiteral $t.TemplateId)'"
    $lines = @(Get-DmfODataAll -Uri $uri -Operation "lines $($t.TemplateId)" -Headers $authHeaders)
    $linesByTemplate[$t.TemplateId] = $lines
    Write-Info ("  {0,-45} {1,4} line(s)" -f $t.TemplateId, $lines.Count)
}

$resolutionByLabel = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
if (-not $NoResolve) {
    $labels = @($linesByTemplate.Values | ForEach-Object { $_ } | ForEach-Object { [string](Get-DmfProp $_ 'Entity' '') } | Where-Object { $_ } | Select-Object -Unique)
    Write-Step "Resolving $($labels.Count) distinct entities (entity map, then Metadata service)"
    $lineObjs = @($labels | ForEach-Object {
        $known = Get-DmfEntityMapEntry -Map $map -EntityName $_
        [pscustomobject]@{ EntityName = $_; TargetEntity = $(if ($known) { $known.targetEntity } else { $null }) }
    })
    $resolutions = @(Resolve-DmfEntity -Lines $lineObjs -Session $session -Map $map -Language $Language)
    foreach ($r in $resolutions) { $resolutionByLabel[$r.EntityName] = $r }
    $resolvedN   = @($resolutions | Where-Object Status -eq 'Resolved').Count
    $notPublicN  = @($resolutions | Where-Object Status -eq 'NotPublic').Count
    $unresolved  = @($resolutions | Where-Object { $_.Status -like 'Unresolved*' })
    Write-Info "$resolvedN resolved  |  $notPublicN not OData-enabled  |  $($unresolved.Count) unresolved"
    foreach ($u in $unresolved) { Write-Detail "  unresolved: $($u.EntityName) -- $($u.Reason)" }
    [void](Save-DmfEntityMap -Map $map)
    Write-Info "Entity map saved ($($map.Entities.Count) entries)."
}

# =============================================================================
#  8.  Write template folders
# =============================================================================
$results     = [System.Collections.Generic.List[pscustomobject]]::new()
$scriptStart = Get-Date
$capturedAt  = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')

foreach ($t in $selected) {
    $tStart = Get-Date
    $result = [pscustomobject]@{
        TemplateId  = $t.TemplateId
        Description = $t.Description
        Lines       = 0
        Resolved    = 0
        Unresolved  = 0
        Folder      = $null
        Status      = 'Failed'
        Elapsed     = '-'
    }
    Write-Step "Capturing  $($t.TemplateId)"
    try {
        $rawLines = @($linesByTemplate[$t.TemplateId])
        $result.Lines = $rawLines.Count
        if ($rawLines.Count -eq 0) {
            Write-Warn 'No lines -- skipping.'
            $result.Status = 'Skipped-NoLines'
            continue
        }

        $folderName = ConvertTo-DmfSafeFileName -Name $t.TemplateId
        $folder     = Join-Path $ResourcesPath $folderName
        $result.Folder = $folder
        if ((Get-TemplateOrigin -Folder $folder) -eq 'custom') {
            Write-Warn "Folder is marked custom in template.json -- never overwritten by capture (rename the environment template or change the origin): $folder"
            $result.Status = 'Skipped-Custom'
            continue
        }
        if ((Test-Path -LiteralPath (Join-Path $folder 'Manifest.xml') -PathType Leaf) -and -not $Force) {
            Write-Warn "Folder exists -- skipping (use -Force to overwrite): $folder"
            $result.Status = 'Skipped-Exists'
            continue
        }

        $manifestLines = [System.Collections.Generic.List[pscustomobject]]::new()
        $sidecarLines  = [System.Collections.Generic.List[object]]::new()
        $sorted = $rawLines | Sort-Object -Property @{ e = { [int](Get-DmfProp $_ 'ExecutionUnit' 1) } }, @{ e = { [int](Get-DmfProp $_ 'LevelInExecutionUnit' 1) } }, @{ e = { [int](Get-DmfProp $_ 'Sequence' 1) } }
        foreach ($l in $sorted) {
            $label  = [string](Get-DmfProp $l 'Entity' '')
            if (-not $label) { continue }
            $target = $null
            if ($resolutionByLabel.ContainsKey($label)) {
                $target = $resolutionByLabel[$label].TargetEntity
                if ($resolutionByLabel[$label].Status -like 'Unresolved*') { $result.Unresolved++ } else { $result.Resolved++ }
            } else {
                $known = Get-DmfEntityMapEntry -Map $map -EntityName $label
                if ($known) { $target = $known.targetEntity }
            }
            $manifestLines.Add([pscustomobject]@{
                EntityName               = $label
                TargetEntity             = $target
                ExecutionUnit            = [int](Get-DmfProp $l 'ExecutionUnit' 1)
                LevelInExecutionUnit     = [int](Get-DmfProp $l 'LevelInExecutionUnit' 1)
                SequenceInLevel          = [int](Get-DmfProp $l 'Sequence' 1)
                FailLevelOnError         = ((Get-DmfProp $l 'FailLevelOnError' 'No') -eq 'Yes')
                FailExecutionUnitOnError = ((Get-DmfProp $l 'FailExecutionUnitOnError' 'No') -eq 'Yes')
            })
            $sidecarLines.Add([ordered]@{
                entity         = $label
                sysModule      = [string](Get-DmfProp $l 'SysModule' '')
                tags           = [string](Get-DmfProp $l 'Tags' '')
                entityCategory = [string](Get-DmfProp $l 'EntityCategory' '')
            })
        }

        if (-not (Test-Path -LiteralPath $folder -PathType Container)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }

        # A folder that already holds data files is a package: its manifest
        # carries field maps and query data that a captured template does not.
        # Keep the original next to the new one so the import path can be
        # restored (git also has it if the folder is tracked).
        $existingManifest = Join-Path $folder 'Manifest.xml'
        if ((Test-Path -LiteralPath $existingManifest -PathType Leaf) -and
            @(Get-ChildItem -LiteralPath $folder -Filter '*.xlsx' -File).Count -gt 0) {
            $backup = Join-Path $folder "Manifest.$(Get-Date -Format 'yyyyMMdd_HHmmss').package.bak.xml"
            Copy-Item -LiteralPath $existingManifest -Destination $backup -Force
            Write-Warn "Folder contains data files; the previous manifest (with field maps) was kept as $(Split-Path -Leaf $backup).  Restore it before importing this package."
        }

        $doc = New-DmfManifestDocument -DefinitionGroupName $t.TemplateId -Description $t.Description -Lines $manifestLines.ToArray()
        Write-DmfManifest      -Document $doc -Path (Join-Path $folder 'Manifest.xml')
        Write-DmfPackageHeader -Path (Join-Path $folder 'PackageHeader.xml') -Description $t.Description
        if (-not $NoSidecar) {
            Write-TemplateSidecar -Folder $folder -Sidecar ([ordered]@{
                schemaVersion = 1
                templateId    = $t.TemplateId
                origin        = 'captured'
                description   = $t.Description
                capturedFrom  = $baseUrl
                capturedAt    = $capturedAt
                capturedBy    = "Export-TemplateDefinition.ps1 $($Script:Version)"
                lines         = $sidecarLines.ToArray()
            })
        }
        Write-Info "$($manifestLines.Count) line(s) written to $folder"
        if ($result.Unresolved -gt 0) { Write-Warn "$($result.Unresolved) line(s) have no TargetEntity yet; -Mode OData will try to resolve them at run time." }
        $result.Status = 'Captured'
    }
    catch {
        $result.Status = 'Failed'
        Write-Fail "Template '$($t.TemplateId)' failed: $_"
        Write-Verbose $_.ScriptStackTrace
    }
    finally {
        $result.Elapsed = Format-Elapsed ((Get-Date) - $tStart)
        $results.Add($result)
        if ($PassThru) { Write-Output $result }
    }
}

# =============================================================================
#  9.  Summary
# =============================================================================
Write-Host ''
Write-Rule "Summary  --  $($results.Count) template(s)  |  $(Format-Elapsed ((Get-Date) - $scriptStart))"
Write-Host ("  {0,-16} {1,6} {2,9} {3,11}  {4}" -f 'Status', 'Lines', 'Resolved', 'Unresolved', 'Template') -ForegroundColor White
foreach ($r in $results) {
    $colour = switch -Wildcard ($r.Status) { 'Captured' { 'Green' } 'Skipped*' { 'DarkGray' } default { 'Red' } }
    Write-Host ("  {0,-16} {1,6} {2,9} {3,11}  {4}" -f $r.Status, $r.Lines, $r.Resolved, $r.Unresolved, $r.TemplateId) -ForegroundColor $colour
}
Write-Host ''
Stop-RunTranscript
if (@($results | Where-Object Status -eq 'Failed').Count -gt 0) { exit 1 }
