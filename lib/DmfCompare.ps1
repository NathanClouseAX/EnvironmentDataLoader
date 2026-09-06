<#
.SYNOPSIS
    Comparison engine for OData snapshot folders (data/<env>/<LE>/) -- pure
    functions so the algorithm is unit-testable without files or a network.

.DESCRIPTION
    Dot-source this file (after DmfOutput.ps1, DmfTemplate.ps1 and DmfPull.ps1)
    to import:

        Get-DmfCompareDefaultIgnores   -- the default ignored-field patterns
        Test-DmfIgnoredField           -- wildcard match against the patterns
        ConvertTo-DmfComparableValue   -- value normalisation (numbers, dates, booleans, blanks)
        Get-DmfCompareKey              -- composite key string for a record
        Compare-DmfEntityRecords       -- Added / Removed / Changed / drift for two record sets
        Compare-DmfSnapshotFolders     -- the same for every entity in two folders
        ConvertTo-DmfCompareFindings   -- flat, CSV-friendly finding objects

    Terminology follows Compare-Object: Reference is the baseline, Difference
    the side under examination.  Added = only in Difference; Removed = only in
    Reference.

    Algorithm
    ─────────
      1. Entities present on one side only are reported with the reason from
         that side's _pull.json (not OData-enabled, failed, not in scope).
      2. Records are keyed by the entity's OData key fields minus any ignored
         field; with no key left, every common non-ignored field forms the key
         and only Added / Removed can be detected.  Duplicate keys are reported.
      3. Records present on both sides are compared field by field over the
         fields both sides have; a field on one side only is schema drift,
         reported once per entity.  Values are normalised unless -Strict.
#>

$Script:DmfCompareDefaultIgnores = @(
    '@odata.etag',
    'CreatedDateTime', 'CreatedBy', 'CreatedTransactionId',
    'ModifiedDateTime', 'ModifiedBy', 'ModifiedTransactionId',
    'RecId', 'RecVersion', 'Partition', '*RecId'
)
$Script:DmfCompareKeySeparator = [string][char]0x1F


function Get-DmfCompareDefaultIgnores {
    return @($Script:DmfCompareDefaultIgnores)
}


function Test-DmfIgnoredField {
    param([Parameter(Mandatory)][string]$Name, [AllowEmptyCollection()][string[]]$Patterns)
    foreach ($p in $Patterns) { if ($Name -like $p) { return $true } }
    return $false
}


function ConvertTo-DmfComparableValue {
    <#
    .SYNOPSIS
        Normalises a field value so equivalent representations compare equal.

    .DESCRIPTION
        Without -Strict:
          - $null, '' and whitespace-only  -> $null
          - booleans and 'true/false/yes/no' strings (any case) -> 'true' / 'false'
          - numbers -> shortest invariant decimal text; strings with a fraction or
            exponent ('1.50', '1E3') are normalised the same way, but integer-looking
            strings are left alone so codes such as '0010' keep their leading zeros
          - dates (typed, or ISO-8601 strings) -> UTC round-trip text; the D365
            "no date" sentinel 1900-01-01T00:00:00Z -> $null
          - arrays / nested objects -> compact JSON
          - other strings -> trimmed, case preserved
        With -Strict only $null is returned for $null; everything else is the raw
        string form.
    #>
    param([AllowNull()]$Value, [switch]$Strict)

    if ($null -eq $Value) { return $null }
    if ($Value -is [System.DBNull]) { return $null }

    if ($Strict) {
        if ($Value -is [string]) { return $Value }
        if ($Value -is [array] -or $Value -is [pscustomobject]) { return ($Value | ConvertTo-Json -Compress -Depth 8) }
        return [string]$Value
    }

    $inv = [System.Globalization.CultureInfo]::InvariantCulture

    if ($Value -is [bool]) { return $(if ($Value) { 'true' } else { 'false' }) }

    if ($Value -is [DateTime]) {
        $utc = if ($Value.Kind -eq [System.DateTimeKind]::Unspecified) { [DateTime]::SpecifyKind($Value, [System.DateTimeKind]::Utc) } else { $Value.ToUniversalTime() }
        if ($utc.Year -eq 1900 -and $utc.Month -eq 1 -and $utc.Day -eq 1 -and $utc.TimeOfDay.Ticks -eq 0) { return $null }
        return $utc.ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ', $inv)
    }
    if ($Value -is [DateTimeOffset]) {
        $utc = $Value.UtcDateTime
        if ($utc.Year -eq 1900 -and $utc.Month -eq 1 -and $utc.Day -eq 1 -and $utc.TimeOfDay.Ticks -eq 0) { return $null }
        return $utc.ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ', $inv)
    }

    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int] -or $Value -is [long] -or
        $Value -is [single] -or $Value -is [double] -or $Value -is [decimal] -or $Value -is [uint16] -or $Value -is [uint32] -or $Value -is [uint64]) {
        try { return ([decimal]$Value).ToString('G29', $inv) } catch { return $Value.ToString($inv) }
    }

    if ($Value -is [string]) {
        $s = $Value.Trim()
        if ($s -eq '') { return $null }
        switch ($s.ToLowerInvariant()) {
            'true'  { return 'true' }
            'false' { return 'false' }
            'yes'   { return 'true' }
            'no'    { return 'false' }
        }
        # fraction / exponent -> normalise; plain integers are left as text
        if ($s -match '^[-+]?(\d+\.\d*|\.\d+|\d+)([eE][-+]?\d+)?$' -and ($s -match '[.eE]')) {
            $d = 0
            if ([decimal]::TryParse($s, [System.Globalization.NumberStyles]::Float, $inv, [ref]$d)) { return $d.ToString('G29', $inv) }
        }
        if ($s -match '^\d{4}-\d{2}-\d{2}(T|$)') {
            $dto = [DateTimeOffset]::MinValue
            $styles = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal
            if ([DateTimeOffset]::TryParse($s, $inv, $styles, [ref]$dto)) {
                $utc = $dto.UtcDateTime
                if ($utc.Year -eq 1900 -and $utc.Month -eq 1 -and $utc.Day -eq 1 -and $utc.TimeOfDay.Ticks -eq 0) { return $null }
                return $utc.ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ', $inv)
            }
        }
        return $s
    }

    if ($Value -is [array] -or $Value -is [System.Collections.IList] -or $Value -is [pscustomobject] -or $Value -is [hashtable]) {
        return ($Value | ConvertTo-Json -Compress -Depth 8)
    }

    return [string]$Value
}


function Get-DmfCompareKey {
    <#
    .SYNOPSIS  Composite key string (normalised values joined with U+001F) for a record.
    #>
    param(
        [Parameter(Mandatory)]$Record,
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$KeyFields,
        [switch]$Strict
    )
    $parts = foreach ($k in $KeyFields) {
        $p = $Record.PSObject.Properties[$k]
        $v = if ($null -ne $p) { ConvertTo-DmfComparableValue -Value $p.Value -Strict:$Strict } else { $null }
        if ($null -eq $v) { '' } else { [string]$v }
    }
    return (@($parts) -join $Script:DmfCompareKeySeparator)
}


function Get-DmfRecordFieldNames {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Records)
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    $list = [System.Collections.Generic.List[string]]::new()
    foreach ($r in $Records) { foreach ($p in $r.PSObject.Properties) { if ($seen.Add($p.Name)) { $list.Add($p.Name) } } }
    return $list.ToArray()
}


function Compare-DmfEntityRecords {
    <#
    .SYNOPSIS
        Compares two record sets of one entity.

    .PARAMETER KeyFields
        Record identity.  Ignored fields are removed from the key; if nothing
        remains the comparison is keyless (every common non-ignored field forms
        the key, so only Added / Removed can be detected).

    .PARAMETER ReferenceFields / DifferenceFields
        Field lists from the snapshot envelopes; derived from the records when
        omitted.  Used for schema drift and to avoid flagging a field that one
        side never had as a per-record change.

    .OUTPUTS
        [pscustomobject]  KeyFields, Keyless, ReferenceCount, DifferenceCount,
        Added[], Removed[], Changed[] (Key, KeyValues, Fields[] of Field /
        ReferenceValue / DifferenceValue), UnchangedCount, KeyCollisions
        (Reference[], Difference[]), SchemaDrift (ReferenceOnly[], DifferenceOnly[]),
        IgnoredFields[]
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Reference,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Difference,
        [AllowEmptyCollection()][string[]]$KeyFields = @(),
        [AllowEmptyCollection()][string[]]$IgnorePatterns,
        [AllowEmptyCollection()][string[]]$ReferenceFields,
        [AllowEmptyCollection()][string[]]$DifferenceFields,
        [switch]$Strict
    )

    if (-not $PSBoundParameters.ContainsKey('IgnorePatterns')) { $IgnorePatterns = Get-DmfCompareDefaultIgnores }
    if ($null -eq $IgnorePatterns) { $IgnorePatterns = @() }

    # @() around each if: an empty result would otherwise collapse to $null
    # (and a one-element result to a scalar) when an if is used as an expression.
    $refFields  = @(if ($ReferenceFields  -and $ReferenceFields.Count  -gt 0) { $ReferenceFields }  else { Get-DmfRecordFieldNames -Records $Reference })
    $diffFields = @(if ($DifferenceFields -and $DifferenceFields.Count -gt 0) { $DifferenceFields } else { Get-DmfRecordFieldNames -Records $Difference })

    $refSet  = [System.Collections.Generic.HashSet[string]]::new([string[]]$refFields,  [System.StringComparer]::OrdinalIgnoreCase)
    $diffSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$diffFields, [System.StringComparer]::OrdinalIgnoreCase)

    $ignoredSeen = [System.Collections.Generic.List[string]]::new()
    $common      = [System.Collections.Generic.List[string]]::new()
    $refOnly     = [System.Collections.Generic.List[string]]::new()
    $diffOnly    = [System.Collections.Generic.List[string]]::new()
    foreach ($f in @($refFields + $diffFields | Select-Object -Unique)) {
        if (Test-DmfIgnoredField -Name $f -Patterns $IgnorePatterns) { $ignoredSeen.Add($f); continue }
        $inRef = $refSet.Contains($f); $inDiff = $diffSet.Contains($f)
        if ($inRef -and $inDiff) { $common.Add($f) }
        elseif ($inRef)          { $refOnly.Add($f) }
        else                     { $diffOnly.Add($f) }
    }

    $effectiveKey = @($KeyFields | Where-Object { $_ -and -not (Test-DmfIgnoredField -Name $_ -Patterns $IgnorePatterns) })
    $keyless = ($effectiveKey.Count -eq 0)
    $keyUsed = @(if ($keyless) { $common } else { $effectiveKey })

    $build = {
        param($records)
        $dict = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $collisions = [System.Collections.Generic.List[string]]::new()
        foreach ($r in $records) {
            $k = Get-DmfCompareKey -Record $r -KeyFields $keyUsed -Strict:$Strict
            if ($dict.ContainsKey($k)) { $collisions.Add($k); continue }
            $dict[$k] = $r
        }
        return @{ Dict = $dict; Collisions = $collisions }
    }
    $ref  = & $build $Reference
    $diff = & $build $Difference

    $added     = [System.Collections.Generic.List[object]]::new()
    $removed   = [System.Collections.Generic.List[object]]::new()
    $changed   = [System.Collections.Generic.List[object]]::new()
    $unchanged = 0

    foreach ($k in $ref.Dict.Keys) {
        if (-not $diff.Dict.ContainsKey($k)) { $removed.Add($ref.Dict[$k]); continue }
        if ($keyless) { $unchanged++; continue }

        $a = $ref.Dict[$k]; $b = $diff.Dict[$k]
        $fieldChanges = [System.Collections.Generic.List[object]]::new()
        foreach ($f in $common) {
            if ($effectiveKey -contains $f) { continue }
            $pa = $a.PSObject.Properties[$f]; $pb = $b.PSObject.Properties[$f]
            $va = ConvertTo-DmfComparableValue -Value $(if ($null -ne $pa) { $pa.Value } else { $null }) -Strict:$Strict
            $vb = ConvertTo-DmfComparableValue -Value $(if ($null -ne $pb) { $pb.Value } else { $null }) -Strict:$Strict
            $equal = if ($null -eq $va -and $null -eq $vb) { $true }
                     elseif ($null -eq $va -or $null -eq $vb) { $false }
                     else { [string]::Equals([string]$va, [string]$vb, [System.StringComparison]::Ordinal) }
            if (-not $equal) {
                $fieldChanges.Add([pscustomobject]@{
                    Field           = $f
                    ReferenceValue  = $(if ($null -ne $pa) { $pa.Value } else { $null })
                    DifferenceValue = $(if ($null -ne $pb) { $pb.Value } else { $null })
                })
            }
        }
        if ($fieldChanges.Count -eq 0) { $unchanged++; continue }

        $keyValues = [ordered]@{}
        foreach ($kf in $keyUsed) { $p = $a.PSObject.Properties[$kf]; $keyValues[$kf] = $(if ($null -ne $p) { $p.Value } else { $null }) }
        $changed.Add([pscustomobject]@{
            Key       = $k
            KeyValues = $keyValues
            Fields    = $fieldChanges.ToArray()
            Reference = $a
            Difference = $b
        })
    }
    foreach ($k in $diff.Dict.Keys) {
        if (-not $ref.Dict.ContainsKey($k)) { $added.Add($diff.Dict[$k]) }
    }

    return [pscustomobject]@{
        KeyFields       = $keyUsed
        Keyless         = $keyless
        ReferenceCount  = $Reference.Count
        DifferenceCount = $Difference.Count
        Added           = $added.ToArray()
        Removed         = $removed.ToArray()
        Changed         = $changed.ToArray()
        UnchangedCount  = $unchanged
        KeyCollisions   = [pscustomobject]@{ Reference = $ref.Collisions.ToArray(); Difference = $diff.Collisions.ToArray() }
        SchemaDrift     = [pscustomobject]@{ ReferenceOnly = $refOnly.ToArray(); DifferenceOnly = $diffOnly.ToArray() }
        IgnoredFields   = $ignoredSeen.ToArray()
    }
}


function Get-DmfSnapshotFilesInFolder {
    param([Parameter(Mandatory)][string]$Folder, [AllowEmptyCollection()][string[]]$Entity)
    $files = @(Get-ChildItem -LiteralPath $Folder -Filter '*.json' -File | Where-Object { $_.Name -ne '_pull.json' })
    if ($Entity -and $Entity.Count -gt 0) {
        $files = @($files | Where-Object {
            $base = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
            $hit = $false
            foreach ($pat in $Entity) { if ($base -like $pat) { $hit = $true; break } }
            $hit
        })
    }
    return $files
}


function Compare-DmfSnapshotFolders {
    <#
    .SYNOPSIS
        Compares every entity snapshot in two data/<env>/<LE>/ folders.

    .PARAMETER Entity
        Wildcard patterns matched against the entity label / file name.

    .PARAMETER KeyOverride
        Hashtable @{ 'Entity label' = 'Field1','Field2' } replacing stored keys.

    .OUTPUTS
        [pscustomobject]  Reference (Path, Index), Difference (Path, Index),
        IgnorePatterns, Entities[] each with: Entity, Status
        (Compared | OnlyInReference | OnlyInDifference), Reason, Reference
        (snapshot or $null), Difference (snapshot or $null), Result
        (Compare-DmfEntityRecords output or $null), KeyFieldsDiffer, Truncated
    #>
    param(
        [Parameter(Mandatory)][string]$ReferencePath,
        [Parameter(Mandatory)][string]$DifferencePath,
        [AllowEmptyCollection()][string[]]$Entity,
        [AllowEmptyCollection()][string[]]$Template,
        [AllowEmptyCollection()][string[]]$IgnorePatterns,
        [hashtable]$KeyOverride,
        [switch]$Strict
    )

    if (-not $PSBoundParameters.ContainsKey('IgnorePatterns')) { $IgnorePatterns = Get-DmfCompareDefaultIgnores }
    if ($null -eq $IgnorePatterns) { $IgnorePatterns = @() }
    foreach ($p in $ReferencePath, $DifferencePath) {
        if (-not (Test-Path -LiteralPath $p -PathType Container)) { throw "Snapshot folder not found: '$p'" }
    }

    $refIndex  = Read-DmfPullIndex -Folder $ReferencePath
    $diffIndex = Read-DmfPullIndex -Folder $DifferencePath

    # -Template: keep only entities that either side's pull index attributes to
    # a matching template (package).  Wildcards allowed.
    $templateScope = $null
    if ($Template -and $Template.Count -gt 0) {
        $templateScope = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($index in $refIndex, $diffIndex) {
            foreach ($name in $index.entities.Keys) {
                $e = $index.entities[$name]
                $tpls = if ($e.Contains('templates')) { @($e['templates']) } else { @() }
                foreach ($t in $tpls) { foreach ($pat in $Template) { if ([string]$t -like $pat) { [void]$templateScope.Add($name); break } } }
            }
        }
    }

    $load = {
        param($folder)
        $map = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($f in (Get-DmfSnapshotFilesInFolder -Folder $folder -Entity $Entity)) {
            try {
                $s = Read-DmfEntitySnapshot -Path $f.FullName
                $map[$s.Entity] = $s
            } catch { Write-Warn "Skipping '$($f.FullName)': $($_.Exception.Message)" }
        }
        return $map
    }
    $refSnaps  = & $load $ReferencePath
    $diffSnaps = & $load $DifferencePath

    $names = @($refSnaps.Keys + $diffSnaps.Keys | Select-Object -Unique | Sort-Object)
    if ($null -ne $templateScope) { $names = @($names | Where-Object { $templateScope.Contains($_) }) }
    $indexReason = {
        param($index, $name)
        if ($index.entities.ContainsKey($name)) {
            $e = $index.entities[$name]
            $st = if ($e.Contains('status')) { [string]$e['status'] } else { '' }
            $rs = if ($e.Contains('reason')) { [string]$e['reason'] } else { '' }
            if ($rs) { return "$st -- $rs" } else { return $st }
        }
        return 'not in scope of that pull'
    }

    $entities = [System.Collections.Generic.List[object]]::new()
    foreach ($name in $names) {
        $inRef = $refSnaps.ContainsKey($name); $inDiff = $diffSnaps.ContainsKey($name)
        $item = [pscustomobject]@{
            Entity          = $name
            Status          = 'Compared'
            Reason          = ''
            Reference       = $(if ($inRef)  { $refSnaps[$name] }  else { $null })
            Difference      = $(if ($inDiff) { $diffSnaps[$name] } else { $null })
            Result          = $null
            KeyFieldsDiffer = $false
            Truncated       = $false
        }
        if (-not $inDiff) { $item.Status = 'OnlyInReference';  $item.Reason = & $indexReason $diffIndex $name; $entities.Add($item); continue }
        if (-not $inRef)  { $item.Status = 'OnlyInDifference'; $item.Reason = & $indexReason $refIndex  $name; $entities.Add($item); continue }

        $a = $refSnaps[$name]; $b = $diffSnaps[$name]
        $keys = @($a.KeyFields)
        if ($KeyOverride -and $KeyOverride.ContainsKey($name)) { $keys = @($KeyOverride[$name]) }
        elseif ((@($a.KeyFields) -join ',') -ne (@($b.KeyFields) -join ',')) {
            $item.KeyFieldsDiffer = $true
            Write-Warn "'$name': key fields differ between sides ($(@($a.KeyFields) -join ',') vs $(@($b.KeyFields) -join ',')); using the Reference side's."
        }
        $item.Truncated = ($a.Truncated -or $b.Truncated)
        $item.Result = Compare-DmfEntityRecords -Reference $a.Records -Difference $b.Records -KeyFields $keys `
            -IgnorePatterns $IgnorePatterns -ReferenceFields $a.Fields -DifferenceFields $b.Fields -Strict:$Strict
        $entities.Add($item)
    }

    $comparison = [pscustomobject]@{
        Reference      = [pscustomobject]@{ Path = $ReferencePath;  Index = $refIndex }
        Difference     = [pscustomobject]@{ Path = $DifferencePath; Index = $diffIndex }
        IgnorePatterns = @($IgnorePatterns)
        TemplateFilter = @($Template)
        Entities       = $entities.ToArray()
        Templates      = @()
    }
    $comparison.Templates = @(Get-DmfCompareTemplateSummary -Comparison $comparison)
    return $comparison
}


function Get-DmfCompareTemplateSummary {
    <#
    .SYNOPSIS
        Rolls the per-entity results up to the templates (packages) that the
        two pull indexes attribute them to.
    .DESCRIPTION
        An entity belongs to every template listed for it on either side.  A
        template that appears in only one index is still reported, so a package
        pulled on one side but never on the other shows up as such.
    .OUTPUTS
        Objects: Template, Entities (names), EntityCount, Compared, Identical,
        WithChanges, Added, Removed, Changed, OnlyInReference, OnlyInDifference,
        InReferenceIndex, InDifferenceIndex
    #>
    param([Parameter(Mandatory)]$Comparison)

    $byEntity = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($e in $Comparison.Entities) { $byEntity[$e.Entity] = $e }

    # template -> set of entity names, and which side's index knows the template
    $members = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $sides   = @{ Reference = $Comparison.Reference.Index; Difference = $Comparison.Difference.Index }
    foreach ($side in $sides.Keys) {
        $index = $sides[$side]
        if ($null -eq $index -or $null -eq $index.entities) { continue }
        foreach ($name in $index.entities.Keys) {
            if (-not $byEntity.ContainsKey($name)) { continue }     # filtered out or absent on both sides
            $e = $index.entities[$name]
            $tpls = if ($e.Contains('templates')) { @($e['templates']) } else { @() }
            foreach ($t in $tpls) {
                $t = [string]$t
                if (-not $t) { continue }
                if (-not $members.ContainsKey($t)) {
                    $members[$t] = @{ Names = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase); Reference = $false; Difference = $false }
                }
                [void]$members[$t].Names.Add($name)
                $members[$t][$side] = $true
            }
        }
    }

    foreach ($t in ($members.Keys | Sort-Object)) {
        $m = $members[$t]
        $names = @($m.Names | Sort-Object)
        $s = [ordered]@{ Compared = 0; Identical = 0; WithChanges = 0; Added = 0; Removed = 0; Changed = 0; OnlyInReference = 0; OnlyInDifference = 0 }
        foreach ($n in $names) {
            $e = $byEntity[$n]
            # if/elseif rather than switch: 'continue' inside a switch only leaves
            # the switch, so one-sided entities would fall through into the
            # compared counters.
            if ($e.Status -eq 'OnlyInReference')  { $s.OnlyInReference++;  continue }
            if ($e.Status -eq 'OnlyInDifference') { $s.OnlyInDifference++; continue }
            if ($null -eq $e.Result) { continue }
            $s.Compared++
            $r = $e.Result
            $delta = $r.Added.Count + $r.Removed.Count + $r.Changed.Count
            if ($delta -eq 0) { $s.Identical++ } else { $s.WithChanges++ }
            $s.Added += $r.Added.Count; $s.Removed += $r.Removed.Count; $s.Changed += $r.Changed.Count
        }
        [pscustomobject]@{
            Template          = $t
            Entities          = $names
            EntityCount       = $names.Count
            Compared          = $s.Compared
            Identical         = $s.Identical
            WithChanges       = $s.WithChanges
            Added             = $s.Added
            Removed           = $s.Removed
            Changed           = $s.Changed
            OnlyInReference   = $s.OnlyInReference
            OnlyInDifference  = $s.OnlyInDifference
            InReferenceIndex  = [bool]$m.Reference
            InDifferenceIndex = [bool]$m.Difference
        }
    }
}


function ConvertTo-DmfCompareFindings {
    <#
    .SYNOPSIS
        Flattens a comparison into one object per finding (CSV-friendly).
    .OUTPUTS
        Objects: Entity, ChangeType (Added | Removed | Changed | EntityOnlyInReference |
        EntityOnlyInDifference | SchemaDrift | KeyCollision), Key, Field,
        ReferenceValue, DifferenceValue, Record
    #>
    param([Parameter(Mandatory)]$Comparison)

    $keyText = { param($rec, $fields) (@(foreach ($f in $fields) { $p = $rec.PSObject.Properties[$f]; "$f=$(if ($null -ne $p) { $p.Value })" }) -join '; ') }
    $emit = { param($e, $t, $k, $f, $rv, $dv, $rec)
        [pscustomobject]@{ Entity = $e; ChangeType = $t; Key = $k; Field = $f; ReferenceValue = $rv; DifferenceValue = $dv; Record = $rec }
    }

    foreach ($item in $Comparison.Entities) {
        $e = $item.Entity
        if ($item.Status -eq 'OnlyInReference')  { & $emit $e 'EntityOnlyInReference'  $null $null $null $null $null; continue }
        if ($item.Status -eq 'OnlyInDifference') { & $emit $e 'EntityOnlyInDifference' $null $null $null $null $null; continue }
        $r = $item.Result
        foreach ($f in $r.SchemaDrift.ReferenceOnly)  { & $emit $e 'SchemaDrift' $null $f 'present' $null $null }
        foreach ($f in $r.SchemaDrift.DifferenceOnly) { & $emit $e 'SchemaDrift' $null $f $null 'present' $null }
        foreach ($k in $r.KeyCollisions.Reference)  { & $emit $e 'KeyCollision' ($k -replace [regex]::Escape($Script:DmfCompareKeySeparator), '; ') $null 'duplicate key' $null $null }
        foreach ($k in $r.KeyCollisions.Difference) { & $emit $e 'KeyCollision' ($k -replace [regex]::Escape($Script:DmfCompareKeySeparator), '; ') $null $null 'duplicate key' $null }
        foreach ($rec in $r.Removed) { & $emit $e 'Removed' (& $keyText $rec $r.KeyFields) $null $null $null $rec }
        foreach ($rec in $r.Added)   { & $emit $e 'Added'   (& $keyText $rec $r.KeyFields) $null $null $null $rec }
        foreach ($c in $r.Changed) {
            $k = & $keyText $c.Reference $r.KeyFields
            foreach ($fc in $c.Fields) { & $emit $e 'Changed' $k $fc.Field $fc.ReferenceValue $fc.DifferenceValue $null }
        }
    }
}
