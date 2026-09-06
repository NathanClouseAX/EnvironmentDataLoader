#Requires -Version 5.1
<#
.SYNOPSIS
    Compares two OData snapshot folders (data/<env>/<legal entity>/) produced by
    Invoke-ProjectExport.ps1 -Mode OData and reports what differs.

.DESCRIPTION
    Entirely local -- no sign-in, no API calls.  For every entity present in
    either folder the script reports:

      - entities present on one side only (with the reason from _pull.json:
        not OData-enabled, failed, unresolved, or simply not in scope)
      - records Added (only in Difference), Removed (only in Reference) and
        Changed, identified by the entity's OData key stored in each file
      - for Changed records, exactly which fields moved
      - schema drift: fields that exist on one side only
      - duplicate keys, which usually mean the stored key is wrong

    Values are normalised before comparison unless -Strict: blanks equal null,
    Yes/No equal true/false, 1.50 equals 1.5, ISO dates compare as instants,
    and the D365 "no date" sentinel (1900-01-01) equals null.  Field names in
    -IgnoreFields (plus the defaults: @odata.etag, audit fields, RecId-like
    surrogates) are excluded from both the comparison and the key.

    Terminology follows Compare-Object: Reference is the baseline, Difference
    the side being examined.

    Outputs: a colour-coded console summary, a self-contained HTML report
    (-HtmlPath), optional machine-readable delta (-JsonPath), and with
    -PassThru one object per finding for Export-Csv.

    Library files (in ./lib/)
    ─────────────────────────
    DmfOutput.ps1    -- Write-* helpers, Format-Elapsed, Stop-RunTranscript
    DmfTemplate.ps1  -- ConvertTo-DmfSafeFileName
    DmfPull.ps1      -- Read-DmfEntitySnapshot, Read-DmfPullIndex
    DmfCompare.ps1   -- comparison engine
    DmfHtml.ps1      -- report styling

.PARAMETER ReferencePath
    Full path of the baseline folder, e.g. C:\DMF\contoso-prod\USMF.
    Aliases: -Folder1, -Path1, -Baseline.

.PARAMETER DifferencePath
    Full path of the folder being examined, e.g. C:\DMF\contoso-uat\USMF.
    Aliases: -Folder2, -Path2, -Target.  Added = only here; Removed = only in
    the baseline.

.PARAMETER Reference
    Shorthand '<env>/<LE>' resolved under -DataPath instead of -ReferencePath.

.PARAMETER Difference
    Shorthand '<env>/<LE>' resolved under -DataPath instead of -DifferencePath.

.PARAMETER DataPath
    Root for the shorthand form.  Default: ./data

.PARAMETER Entity
    One or more entity labels; wildcards allowed.  Default: every entity.

.PARAMETER Template
    One or more template (package) names; wildcards allowed.  Keeps only the
    entities that either side's _pull.json attributes to a matching template,
    so you can ask "how does package X differ between these environments".
    The report also rolls every result up per template regardless.

.PARAMETER IgnoreFields
    Field names or wildcard patterns added to the default ignore list.

.PARAMETER NoDefaultIgnores
    Compare every field, including @odata.etag and the audit fields.

.PARAMETER KeyOverride
    Hashtable @{ 'Entity label' = 'Field1','Field2' } replacing the stored key
    for specific entities.

.PARAMETER Strict
    Disable value normalisation; compare raw text.

.PARAMETER ChangesOnly
    Omit identical entities from the console table and the HTML.

.PARAMETER MaxRowsPerEntity
    Cap on rows rendered per entity in the HTML (console and -PassThru are
    never capped).  Default 500.

.PARAMETER HtmlPath
    HTML report path.  Auto-generated under $env:TEMP when omitted; '' suppresses.

.PARAMETER JsonPath
    Write the machine-readable delta: per entity the key fields, counts, schema
    drift, and the full added / removed records and changed fields, plus the
    per-template roll-up.  Intended for tooling that acts on the differences.

.PARAMETER LogPath
    Transcript path; auto-generated when omitted; '' suppresses.

.PARAMETER PassThru
    Emit one object per finding: Entity, ChangeType, Key, Field,
    ReferenceValue, DifferenceValue, Record.

.PARAMETER FailOnDifference
    Exit with code 2 when any difference is found (for CI).

.EXAMPLE
    # Two environments pulled to C:\DMF, compared package by package
    .\Compare-EnvironmentData.ps1 -Folder1 'C:\DMF\contoso-prod\USMF' -Folder2 'C:\DMF\contoso-uat\USMF'

.EXAMPLE
    # Only the entities of one package
    .\Compare-EnvironmentData.ps1 -Folder1 'C:\DMF\contoso-prod\USMF' -Folder2 'C:\DMF\contoso-uat\USMF' -Template '010 - System Setup'

.EXAMPLE
    .\Compare-EnvironmentData.ps1 -Reference 'contoso-uat/USMF' -Difference 'contoso-prod/USMF'

.EXAMPLE
    .\Compare-EnvironmentData.ps1 -ReferencePath 'data\contoso-uat\USMF' -DifferencePath 'data\contoso-prod\USMF' `
        -Entity 'Currencies', 'Number sequence*' -IgnoreFields 'Description' -PassThru |
        Export-Csv delta.csv -NoTypeInformation

.EXAMPLE
    # Cross-company inside one environment
    .\Compare-EnvironmentData.ps1 -Reference 'contoso-uat/USMF' -Difference 'contoso-uat/DAT' -IgnoreFields dataAreaId
#>
[CmdletBinding(DefaultParameterSetName = 'Paths')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Paths')]
    [Alias('Folder1', 'Path1', 'Baseline')]
    [string]$ReferencePath,

    [Parameter(Mandatory, ParameterSetName = 'Paths')]
    [Alias('Folder2', 'Path2', 'Target')]
    [string]$DifferencePath,

    [Parameter(Mandatory, ParameterSetName = 'Short')]
    [string]$Reference,

    [Parameter(Mandatory, ParameterSetName = 'Short')]
    [string]$Difference,

    [string]$DataPath,

    [string[]]$Entity,

    [string[]]$Template,

    [string[]]$IgnoreFields,
    [switch]$NoDefaultIgnores,

    [hashtable]$KeyOverride,

    [switch]$Strict,
    [switch]$ChangesOnly,

    [ValidateRange(1, 100000)]
    [int]$MaxRowsPerEntity = 500,

    [AllowEmptyString()]
    [string]$HtmlPath,

    [string]$JsonPath,

    [AllowEmptyString()]
    [string]$LogPath,

    [switch]$PassThru,
    [switch]$FailOnDifference
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
. (Join-Path $libPath 'DmfTemplate.ps1')
. (Join-Path $libPath 'DmfPull.ps1')
. (Join-Path $libPath 'DmfCompare.ps1')
. (Join-Path $libPath 'DmfHtml.ps1')

# =============================================================================
#  Pre-flight
# =============================================================================
if (-not $DataPath) { $DataPath = Join-Path $PSScriptRoot 'data' }
if ($PSCmdlet.ParameterSetName -eq 'Short') {
    $ReferencePath  = Join-Path $DataPath ($Reference  -replace '/', '\')
    $DifferencePath = Join-Path $DataPath ($Difference -replace '/', '\')
}
foreach ($p in $ReferencePath, $DifferencePath) {
    if (-not (Test-Path -LiteralPath $p -PathType Container)) { throw "Snapshot folder not found: '$p'" }
    if (@(Get-ChildItem -LiteralPath $p -Filter '*.json' -File | Where-Object Name -ne '_pull.json').Count -eq 0) {
        throw "No entity snapshots (*.json) in '$p'.  Pull data first with Invoke-ProjectExport.ps1 -Mode OData."
    }
}
$ReferencePath  = (Resolve-Path -LiteralPath $ReferencePath).Path
$DifferencePath = (Resolve-Path -LiteralPath $DifferencePath).Path

$ignorePatterns = @()
if (-not $NoDefaultIgnores) { $ignorePatterns += Get-DmfCompareDefaultIgnores }
if ($IgnoreFields) { $ignorePatterns += $IgnoreFields }
$ignorePatterns = @($ignorePatterns | Select-Object -Unique)

$Script:Version          = '1.0'
$Script:LineWidth        = 80
$Script:TranscriptActive = $false
try {
    $w = $Host.UI.RawUI.BufferSize.Width
    $Script:LineWidth = [Math]::Min(120, [Math]::Max(80, $w))
} catch { <# non-interactive host -- keep default #> }

$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
if (-not $PSBoundParameters.ContainsKey('LogPath'))  { $LogPath  = Join-Path $env:TEMP "DMFDataDiff_${ts}.log" }
if (-not $PSBoundParameters.ContainsKey('HtmlPath')) { $HtmlPath = Join-Path $env:TEMP "DMFDataDiff_${ts}.html" }
if ($LogPath -ne '') {
    try { Start-Transcript -Path $LogPath -Force | Out-Null; $Script:TranscriptActive = $true } catch {}
}

# =============================================================================
#  1.  Compare
# =============================================================================
$scriptStart = Get-Date
Write-Banner -Title 'D365 F&O Environment Data Diff'
Write-Host ''
Write-Info "Reference  : $ReferencePath"
Write-Info "Difference : $DifferencePath"
Write-Info "Ignored    : $($ignorePatterns -join ', ')$(if ($Strict) { '   [strict values]' })"
if ($Entity)   { Write-Info "Entities   : $($Entity -join ', ')" }
if ($Template) { Write-Info "Templates  : $($Template -join ', ')" }
if ($Script:TranscriptActive) { Write-Info "Log        : $LogPath" }

Write-Step 'Comparing snapshots'
$cmpArgs = @{ ReferencePath = $ReferencePath; DifferencePath = $DifferencePath; IgnorePatterns = $ignorePatterns; Strict = $Strict }
if ($Entity)      { $cmpArgs['Entity']      = $Entity }
if ($Template)    { $cmpArgs['Template']    = $Template }
if ($KeyOverride) { $cmpArgs['KeyOverride'] = $KeyOverride }
$comparison = Compare-DmfSnapshotFolders @cmpArgs

$sideInfo = {
    param($side)
    $idx = $side.Index
    [pscustomobject]@{
        Path        = $side.Path
        Environment = $(if ($idx.environment) { [string]$idx.environment } else { Split-Path (Split-Path $side.Path -Parent) -Leaf })
        Url         = $(if ($idx.environmentUrl) { [string]$idx.environmentUrl } else { '' })
        LegalEntity = $(if ($idx.legalEntity) { [string]$idx.legalEntity } else { Split-Path $side.Path -Leaf })
        PulledAt    = $(if ($null -ne $idx.lastRun -and $null -ne $idx.lastRun.PSObject.Properties['finishedAt']) { [string]$idx.lastRun.finishedAt } else { '' })
        Templates   = $(if ($null -ne $idx.lastRun -and $null -ne $idx.lastRun.PSObject.Properties['templates']) { @($idx.lastRun.templates) -join ', ' } else { '' })
    }
}
$refInfo  = & $sideInfo $comparison.Reference
$diffInfo = & $sideInfo $comparison.Difference

# =============================================================================
#  2.  Totals
# =============================================================================
$rows = [System.Collections.Generic.List[pscustomobject]]::new()
$tot  = [ordered]@{ Entities = 0; Compared = 0; Identical = 0; WithChanges = 0; Added = 0; Removed = 0; Changed = 0; Drift = 0; Warnings = 0; OnlyOneSide = 0; Truncated = 0 }
foreach ($e in $comparison.Entities) {
    $tot.Entities++
    $row = [pscustomobject]@{
        Entity = $e.Entity; Status = $e.Status; Ref = '-'; Diff = '-'; Added = '-'; Removed = '-'; Changed = '-'
        Notes = @(); Level = 'ok'; Identical = $false; Item = $e
    }
    if ($e.Status -ne 'Compared') {
        $tot.OnlyOneSide++
        $side = if ($e.Status -eq 'OnlyInReference') { 'Difference' } else { 'Reference' }
        $row.Notes += "only in $(if ($e.Status -eq 'OnlyInReference') { 'Reference' } else { 'Difference' }) ($side side: $($e.Reason))"
        $row.Level = if ($e.Reason -like 'NotPublic*' -or $e.Reason -like 'not in scope*') { 'mute' } else { 'err' }
        if ($row.Level -eq 'err') { $tot.Warnings++ }
        $rows.Add($row); continue
    }
    $r = $e.Result
    $tot.Compared++
    $row.Ref = $r.ReferenceCount; $row.Diff = $r.DifferenceCount
    $row.Added = $r.Added.Count; $row.Removed = $r.Removed.Count; $row.Changed = $r.Changed.Count
    $tot.Added += $r.Added.Count; $tot.Removed += $r.Removed.Count; $tot.Changed += $r.Changed.Count
    $drift = $r.SchemaDrift.ReferenceOnly.Count + $r.SchemaDrift.DifferenceOnly.Count
    if ($drift -gt 0) { $tot.Drift += $drift; $row.Notes += "SchemaDrift($drift)" }
    $coll = $r.KeyCollisions.Reference.Count + $r.KeyCollisions.Difference.Count
    if ($coll -gt 0) { $tot.Warnings++; $row.Notes += "KeyCollision($coll)"; $row.Level = 'err' }
    if ($r.Keyless) { $tot.Warnings++; $row.Notes += 'KeylessComparison'; if ($row.Level -eq 'ok') { $row.Level = 'warn' } }
    if ($e.KeyFieldsDiffer) { $row.Notes += 'key fields differ' }
    if ($e.Truncated) { $tot.Truncated++; $row.Notes += 'Truncated'; if ($row.Level -eq 'ok') { $row.Level = 'warn' } }
    $hasChanges = ($r.Added.Count + $r.Removed.Count + $r.Changed.Count) -gt 0
    if ($hasChanges) { $tot.WithChanges++; if ($row.Level -eq 'ok') { $row.Level = 'warn' } }
    else { $row.Identical = $true; if ($drift -eq 0) { $tot.Identical++ } else { $tot.Identical++ } }
    $rows.Add($row)
}
$anyDifference = ($tot.Added + $tot.Removed + $tot.Changed + $tot.Drift + $tot.OnlyOneSide) -gt 0

# =============================================================================
#  3.  Console
# =============================================================================
Write-Host ''
Write-Rule "Result  --  $($tot.Entities) entities  |  $(Format-Elapsed ((Get-Date) - $scriptStart))"
$fmt = '  {0,-42} {1,8} {2,8} {3,7} {4,8} {5,8}  {6}'
Write-Host ($fmt -f 'Entity', 'Ref', 'Diff', 'Added', 'Removed', 'Changed', 'Notes') -ForegroundColor White
Write-Host ($fmt -f ('-' * 42), ('-' * 8), ('-' * 8), ('-' * 7), ('-' * 8), ('-' * 8), ('-' * 20)) -ForegroundColor DarkGray
$suppressed = 0
foreach ($row in $rows) {
    if ($ChangesOnly -and $row.Identical -and $row.Notes.Count -eq 0) { $suppressed++; continue }
    $colour = switch ($row.Level) { 'ok' { 'Green' } 'warn' { 'Yellow' } 'err' { 'Red' } default { 'DarkGray' } }
    $name = if ($row.Entity.Length -gt 42) { $row.Entity.Substring(0, 39) + '...' } else { $row.Entity }
    Write-Host ($fmt -f $name, $row.Ref, $row.Diff, $row.Added, $row.Removed, $row.Changed, ($row.Notes -join '; ')) -ForegroundColor $colour
}
if ($suppressed -gt 0) { Write-Host "  ($suppressed identical entities hidden by -ChangesOnly)" -ForegroundColor DarkGray }
Write-Host ''
$summaryLine = "  $($tot.Compared) compared: $($tot.Identical) identical, $($tot.WithChanges) with changes  |  rows: +$($tot.Added) -$($tot.Removed) ~$($tot.Changed)  |  schema drift: $($tot.Drift)  |  one side only: $($tot.OnlyOneSide)  |  warnings: $($tot.Warnings)"
Write-Host $summaryLine -ForegroundColor $(if (-not $anyDifference) { 'Green' } elseif ($tot.Warnings -gt 0) { 'Red' } else { 'Yellow' })

# -- Per-template (package) roll-up ------------------------------------------
$templates = @($comparison.Templates)
if ($templates.Count -gt 0) {
    Write-Host ''
    Write-Rule 'By template (package)'
    $tfmt = '  {0,-38} {1,8} {2,9} {3,8} {4,7} {5,8} {6,8}  {7}'
    Write-Host ($tfmt -f 'Template', 'Entities', 'Identical', 'Changed', 'Added', 'Removed', '~Rows', 'Notes') -ForegroundColor White
    Write-Host ($tfmt -f ('-' * 38), ('-' * 8), ('-' * 9), ('-' * 8), ('-' * 7), ('-' * 8), ('-' * 8), ('-' * 12)) -ForegroundColor DarkGray
    foreach ($t in $templates) {
        if ($ChangesOnly -and $t.WithChanges -eq 0 -and ($t.OnlyInReference + $t.OnlyInDifference) -eq 0) { continue }
        $notes = @()
        if (-not $t.InReferenceIndex)  { $notes += 'not pulled on Reference' }
        if (-not $t.InDifferenceIndex) { $notes += 'not pulled on Difference' }
        if ($t.OnlyInReference)  { $notes += "$($t.OnlyInReference) only in Reference" }
        if ($t.OnlyInDifference) { $notes += "$($t.OnlyInDifference) only in Difference" }
        $colour = if ($t.WithChanges -eq 0 -and $notes.Count -eq 0) { 'Green' } elseif ($notes.Count -gt 0) { 'Yellow' } else { 'Yellow' }
        $name = if ($t.Template.Length -gt 38) { $t.Template.Substring(0, 35) + '...' } else { $t.Template }
        Write-Host ($tfmt -f $name, $t.EntityCount, $t.Identical, $t.WithChanges, $t.Added, $t.Removed, $t.Changed, ($notes -join '; ')) -ForegroundColor $colour
    }
}

# =============================================================================
#  4.  HTML report
# =============================================================================
if ($HtmlPath -ne '') {
    $h = { param($t) ConvertTo-HtmlEncoded ([string]$t) }
    $val = { param($v) if ($null -eq $v) { '<span class="kv">(null)</span>' } elseif ($v -is [array] -or $v -is [pscustomobject]) { & $h ($v | ConvertTo-Json -Compress -Depth 5) } else { & $h ([string]$v) } }
    $recordCell = {
        param($rec, $keys, $ignore)
        $parts = foreach ($p in $rec.PSObject.Properties) {
            if (Test-DmfIgnoredField -Name $p.Name -Patterns $ignore) { continue }
            "<b>$(& $h $p.Name)</b>=$(& $val $p.Value)"
        }
        "<span class='kv'>$(@($parts) -join '; ')</span>"
    }
    $keyCell = { param($rec, $keys) @(foreach ($k in $keys) { $p = $rec.PSObject.Properties[$k]; "<b>$(& $h $k)</b>=$(& $val $(if ($null -ne $p) { $p.Value } else { $null }))" }) -join '; ' }

    $generated = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $header = @"
<header class="page-header">
  <h1>D365 F&amp;O &mdash; Environment Data Diff</h1>
  <div class="meta"><span>Generated: $generated</span><span>Ignored fields: $(& $h ($ignorePatterns -join ', '))</span>$(if ($Strict) { '<span>&#9888; Strict values</span>' })$(if ($ChangesOnly) { '<span>Changes only</span>' })</div>
  <div class="sides">
    <div class="side"><b>Reference (baseline)</b>$(& $h $refInfo.Environment) / $(& $h $refInfo.LegalEntity)<br>$(& $h $refInfo.Url)<br>pulled $(& $h $refInfo.PulledAt)$(if ($refInfo.Templates) { " &bull; $(& $h $refInfo.Templates)" })</div>
    <div class="side"><b>Difference (examined)</b>$(& $h $diffInfo.Environment) / $(& $h $diffInfo.LegalEntity)<br>$(& $h $diffInfo.Url)<br>pulled $(& $h $diffInfo.PulledAt)$(if ($diffInfo.Templates) { " &bull; $(& $h $diffInfo.Templates)" })</div>
  </div>
</header>
"@

    $body = [System.Text.StringBuilder]::new()
    [void]$body.AppendLine('<section class="summary">')
    [void]$body.AppendLine("  <div class='card'><div class='val'>$($tot.Compared)</div><div class='lbl'>Entities compared</div></div>")
    [void]$body.AppendLine("  <div class='card ok'><div class='val'>$($tot.Identical)</div><div class='lbl'>Identical</div></div>")
    [void]$body.AppendLine("  <div class='card warn'><div class='val'>$($tot.WithChanges)</div><div class='lbl'>With changes</div></div>")
    [void]$body.AppendLine("  <div class='card add'><div class='val'>$($tot.Added)</div><div class='lbl'>Rows added</div></div>")
    [void]$body.AppendLine("  <div class='card rem'><div class='val'>$($tot.Removed)</div><div class='lbl'>Rows removed</div></div>")
    [void]$body.AppendLine("  <div class='card chg'><div class='val'>$($tot.Changed)</div><div class='lbl'>Rows changed</div></div>")
    [void]$body.AppendLine("  <div class='card'><div class='val'>$($tot.Drift)</div><div class='lbl'>Schema drift</div></div>")
    [void]$body.AppendLine("  <div class='card $(if ($tot.Warnings) { 'err' })'><div class='val'>$($tot.Warnings)</div><div class='lbl'>Warnings</div></div>")
    [void]$body.AppendLine('</section>')
    if ($tot.Truncated -gt 0) { [void]$body.AppendLine("<div class='notice'>&#9888; $($tot.Truncated) entity file(s) were truncated when pulled (-MaxRecordsPerEntity); their counts are not reliable.</div>") }

    # Coverage
    $oneSide = @($rows | Where-Object { $_.Status -ne 'Compared' })
    if ($oneSide.Count -gt 0) {
        [void]$body.AppendLine('<section class="section"><h2>Coverage</h2><div class="sub">Entities present on one side only, with the reason recorded by that side&#39;s pull.</div><div class="table-wrap"><table><thead><tr><th>Entity</th><th>Present in</th><th>Other side</th></tr></thead><tbody>')
        foreach ($row in $oneSide) {
            $cls = if ($row.Level -eq 'err') { 'r-err' } else { 'r-mute' }
            [void]$body.AppendLine("<tr class='$cls'><td>$(& $h $row.Entity)</td><td>$(if ($row.Status -eq 'OnlyInReference') { 'Reference' } else { 'Difference' })</td><td>$(& $h $row.Item.Reason)</td></tr>")
        }
        [void]$body.AppendLine('</tbody></table></div></section>')
    }

    # Per-template (package) roll-up
    if ($templates.Count -gt 0) {
        [void]$body.AppendLine('<section class="section"><h2>By template (package)</h2><div class="sub">Every entity is attributed to the template(s) that pulled it, per each side&#39;s _pull.json.</div><div class="table-wrap"><table><thead><tr><th>Template</th><th class="r">Entities</th><th class="r">Identical</th><th class="r">With changes</th><th class="r">Rows added</th><th class="r">Rows removed</th><th class="r">Rows changed</th><th>Notes</th></tr></thead><tbody>')
        foreach ($t in $templates) {
            $notes = @()
            if (-not $t.InReferenceIndex)  { $notes += 'not pulled on Reference' }
            if (-not $t.InDifferenceIndex) { $notes += 'not pulled on Difference' }
            if ($t.OnlyInReference)  { $notes += "$($t.OnlyInReference) entity(ies) only in Reference" }
            if ($t.OnlyInDifference) { $notes += "$($t.OnlyInDifference) entity(ies) only in Difference" }
            $cls = if ($t.WithChanges -eq 0 -and $notes.Count -eq 0) { 'r-ok' } elseif ($notes.Count -gt 0) { 'r-warn' } else { 'r-warn' }
            [void]$body.AppendLine("<tr class='$cls'><td>$(& $h $t.Template)</td><td class='r'>$($t.EntityCount)</td><td class='r'>$($t.Identical)</td><td class='r'>$($t.WithChanges)</td><td class='r'>$($t.Added)</td><td class='r'>$($t.Removed)</td><td class='r'>$($t.Changed)</td><td>$(& $h ($notes -join '; '))</td></tr>")
        }
        [void]$body.AppendLine('</tbody></table></div></section>')
    }

    # Filter toolbar + per-entity sections
    [void]$body.AppendLine('<div class="toolbar"><label for="flt">Filter</label><input id="flt" type="search" placeholder="entity, key, field or value&hellip;" oninput="dmfFilter(this.value)"><span id="fltinfo" class="kv"></span></div>')
    [void]$body.AppendLine('<section class="section"><h2>Entities</h2></section>')
    foreach ($row in $rows) {
        if ($row.Status -ne 'Compared') { continue }
        if ($ChangesOnly -and $row.Identical -and $row.Notes.Count -eq 0) { continue }
        $r = $row.Item.Result
        $open = if ($row.Identical -and $row.Notes.Count -eq 0) { '' } else { ' open' }
        $badge = if ($row.Level -eq 'err') { 'b-err' } elseif ($row.Identical) { 'b-ok' } else { 'b-warn' }
        $badgeText = if ($row.Level -eq 'err') { 'warning' } elseif ($row.Identical) { 'identical' } else { 'changed' }
        $counts = "ref $($r.ReferenceCount) &bull; diff $($r.DifferenceCount) &bull; <span class='diff-new'>+$($r.Added.Count)</span> / <span class='diff-old'>-$($r.Removed.Count)</span> / ~$($r.Changed.Count)$(if ($row.Notes.Count) { ' &bull; ' + (& $h ($row.Notes -join '; ')) })"
        [void]$body.AppendLine("<details class='entity'$open data-entity='$(& $h $row.Entity)'><summary><span class='badge $badge'>$badgeText</span><span class='name'>$(& $h $row.Entity)</span><span class='counts'>$counts</span></summary><div class='body'>")
        [void]$body.AppendLine("<div class='sub' style='padding:8px 14px 0'>Key: <span class='mono'>$(& $h ($r.KeyFields -join ', '))</span>$(if ($r.Keyless) { ' (keyless: every non-ignored field)' })</div>")
        if ($r.SchemaDrift.ReferenceOnly.Count -or $r.SchemaDrift.DifferenceOnly.Count) {
            [void]$body.AppendLine("<div class='notice info' style='margin:8px 14px'>Schema drift &mdash; only in Reference: <span class='mono'>$(& $h ($r.SchemaDrift.ReferenceOnly -join ', '))</span>; only in Difference: <span class='mono'>$(& $h ($r.SchemaDrift.DifferenceOnly -join ', '))</span></div>")
        }
        if ($r.KeyCollisions.Reference.Count -or $r.KeyCollisions.Difference.Count) {
            [void]$body.AppendLine("<div class='notice' style='margin:8px 14px'>Duplicate keys: $($r.KeyCollisions.Reference.Count) in Reference, $($r.KeyCollisions.Difference.Count) in Difference &mdash; the stored key may be incomplete; consider -KeyOverride.</div>")
        }
        $total = $r.Added.Count + $r.Removed.Count + $r.Changed.Count
        if ($total -gt 0) {
            [void]$body.AppendLine("<div class='table-wrap'><table><thead><tr><th>Change</th><th>Key</th><th>Detail</th></tr></thead><tbody>")
            $rendered = 0
            foreach ($c in $r.Changed) {
                if ($rendered -ge $MaxRowsPerEntity) { break }
                $detail = @(foreach ($fc in $c.Fields) { "<span class='mono'>$(& $h $fc.Field)</span>: <span class='diff-old'>$(& $val $fc.ReferenceValue)</span><span class='arrow'>&rarr;</span><span class='diff-new'>$(& $val $fc.DifferenceValue)</span>" }) -join '<br>'
                [void]$body.AppendLine("<tr class='k-chg'><td><span class='badge b-chg'>changed</span></td><td>$(& $keyCell $c.Reference $r.KeyFields)</td><td>$detail</td></tr>")
                $rendered++
            }
            foreach ($rec in $r.Added) {
                if ($rendered -ge $MaxRowsPerEntity) { break }
                [void]$body.AppendLine("<tr class='k-add'><td><span class='badge b-add'>added</span></td><td>$(& $keyCell $rec $r.KeyFields)</td><td>$(& $recordCell $rec $r.KeyFields $ignorePatterns)</td></tr>")
                $rendered++
            }
            foreach ($rec in $r.Removed) {
                if ($rendered -ge $MaxRowsPerEntity) { break }
                [void]$body.AppendLine("<tr class='k-rem'><td><span class='badge b-rem'>removed</span></td><td>$(& $keyCell $rec $r.KeyFields)</td><td>$(& $recordCell $rec $r.KeyFields $ignorePatterns)</td></tr>")
                $rendered++
            }
            [void]$body.AppendLine('</tbody></table></div>')
            if ($total -gt $rendered) { [void]$body.AppendLine("<div class='more'>&hellip; $($total - $rendered) more row(s) not shown (-MaxRowsPerEntity $MaxRowsPerEntity); use -PassThru or -JsonPath for the full set.</div>") }
        }
        [void]$body.AppendLine('</div></details>')
    }

    $js = @'
function dmfFilter(q){q=(q||'').toLowerCase();var n=0,shown=0;document.querySelectorAll('details.entity').forEach(function(d){n++;var hit=!q||d.textContent.toLowerCase().indexOf(q)>=0;d.style.display=hit?'':'none';if(hit){shown++;if(q){d.open=true;d.querySelectorAll('tbody tr').forEach(function(tr){tr.style.display=tr.textContent.toLowerCase().indexOf(q)>=0||d.getAttribute('data-entity').toLowerCase().indexOf(q)>=0?'':'none';});}else{d.querySelectorAll('tbody tr').forEach(function(tr){tr.style.display='';});}}});var i=document.getElementById('fltinfo');if(i){i.textContent=q?(shown+' of '+n+' entities match'):'';}}
'@
    $footer = "D365 F&amp;O Environment Data Diff &nbsp;&bull;&nbsp; EnvironmentDataLoader / Compare-EnvironmentData.ps1 &nbsp;&bull;&nbsp; $generated"
    $html = New-DmfHtmlDocument -Title "Data diff: $($refInfo.Environment)/$($refInfo.LegalEntity) vs $($diffInfo.Environment)/$($diffInfo.LegalEntity)" -HeaderHtml $header -BodyHtml $body.ToString() -FooterHtml $footer -ScriptJs $js
    [System.IO.File]::WriteAllText($HtmlPath, $html, [System.Text.Encoding]::UTF8)
    Write-OK "HTML report written : $HtmlPath"
}

# =============================================================================
#  5.  JSON delta
# =============================================================================
if ($JsonPath) {
    $delta = [ordered]@{
        schemaVersion = 1
        reference     = [ordered]@{ environment = $refInfo.Environment;  legalEntity = $refInfo.LegalEntity;  path = $ReferencePath;  pulledAt = $refInfo.PulledAt }
        difference    = [ordered]@{ environment = $diffInfo.Environment; legalEntity = $diffInfo.LegalEntity; path = $DifferencePath; pulledAt = $diffInfo.PulledAt }
        ignoredFields = @($ignorePatterns)
        strict        = [bool]$Strict
        templateFilter = @($Template)
        templates     = @(foreach ($t in $templates) {
            [ordered]@{
                template = $t.Template; entities = @($t.Entities)
                counts   = [ordered]@{ entities = $t.EntityCount; compared = $t.Compared; identical = $t.Identical; withChanges = $t.WithChanges; added = $t.Added; removed = $t.Removed; changed = $t.Changed; onlyInReference = $t.OnlyInReference; onlyInDifference = $t.OnlyInDifference }
                pulledOn = [ordered]@{ reference = $t.InReferenceIndex; difference = $t.InDifferenceIndex }
            }
        })
        entities      = [ordered]@{}
    }
    foreach ($e in $comparison.Entities) {
        if ($e.Status -ne 'Compared') {
            $delta.entities[$e.Entity] = [ordered]@{ status = $e.Status; reason = $e.Reason }
            continue
        }
        $r = $e.Result
        $delta.entities[$e.Entity] = [ordered]@{
            status      = 'Compared'
            keyFields   = @($r.KeyFields)
            keyless     = [bool]$r.Keyless
            counts      = [ordered]@{ reference = $r.ReferenceCount; difference = $r.DifferenceCount; added = $r.Added.Count; removed = $r.Removed.Count; changed = $r.Changed.Count; unchanged = $r.UnchangedCount }
            schemaDrift = [ordered]@{ referenceOnly = @($r.SchemaDrift.ReferenceOnly); differenceOnly = @($r.SchemaDrift.DifferenceOnly) }
            keyCollisions = [ordered]@{ reference = $r.KeyCollisions.Reference.Count; difference = $r.KeyCollisions.Difference.Count }
            added       = @($r.Added)
            removed     = @($r.Removed)
            changed     = @(foreach ($c in $r.Changed) {
                $fields = [ordered]@{}
                foreach ($fc in $c.Fields) { $fields[$fc.Field] = [ordered]@{ reference = $fc.ReferenceValue; difference = $fc.DifferenceValue } }
                [ordered]@{ key = $c.KeyValues; fields = $fields }
            })
        }
    }
    $dir = Split-Path -Parent $JsonPath
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [System.IO.File]::WriteAllText($JsonPath, ($delta | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding($false)))
    Write-OK "JSON delta written  : $JsonPath"
}

# =============================================================================
#  6.  Finish
# =============================================================================
Stop-RunTranscript
if ($PassThru) { ConvertTo-DmfCompareFindings -Comparison $comparison | Write-Output }
if ($FailOnDifference -and $anyDifference) { exit 2 }
