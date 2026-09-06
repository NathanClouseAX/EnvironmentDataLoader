<#
.SYNOPSIS
    Entity resolution for the OData path: DMF entity label -> AOT entity name
    -> OData collection, key fields, and company-specific flag.

.DESCRIPTION
    Dot-source this file (after DmfOutput.ps1, DmfRequest.ps1, DmfAuth.ps1,
    DmfOData.ps1 and DmfTemplate.ps1) to import:

        Get-DmfEntityMap / Save-DmfEntityMap     -- resources/entity-map.json cache
        Get-DmfEntityMapEntry, Set-DmfEntityMapEntry, Test-DmfEntityMapResolved
        Update-DmfEntityMapFromManifests         -- harvest label -> TargetEntity from manifests on disk
        Resolve-DmfEntity                        -- resolve template lines (cache first, then Metadata service)
        Get-DmfMetadataDataEntities, Get-DmfMetadataDataEntity,
        Get-DmfMetadataLabels, Find-DmfMetadataLabelIds,
        Get-DmfMetadataPublicEntity              -- thin wrappers over /Metadata/*

    Resolution order
    ────────────────
      1. entity-map.json -- hit and complete -> done (unless -Refresh).
      2. Metadata service:
           - TargetEntity known -> /Metadata/DataEntities?$filter=Name eq '...'
           - else reverse label lookup /Metadata/Labels?$filter=Value eq '<label>'
             -> DataEntities by LabelId.  If the Value filter is not supported
             the whole DataEntities list is fetched once and its labels are
             resolved in Id batches to build a label -> entity index.
           - /Metadata/PublicEntities?$filter=Name eq '<PublicEntityName>'
             -> key fields and whether dataAreaId is present.
      3. Nothing -> Unresolved (the entity is skipped by the pull and reported).

    entity-map.json entries are keyed by the DMF label.  A 'manual' entry is
    never overwritten by resolution (edit the file by hand to pin a mapping).

.NOTES
    All network calls go through Invoke-DmfRequest, so retry, throttling and
    session token refresh apply.
#>

$Script:DmfEntityMapSchemaVersion   = 1
$Script:DmfLabelBatchSize           = 25
$Script:DmfMetadataEntityCache      = @{}     # BaseUrl -> DataEntities rows (per run)
$Script:DmfMetadataPublicCache      = @{}     # BaseUrl -> PublicEntities rows (per run, only if a scan was needed)
$Script:DmfMetadataLabelFilterOk    = $null   # $null = untested; $false once the service rejects $filter on Labels
$Script:DmfMetadataLabelTextCache   = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)  # "lang|id" -> text
$Script:DmfMetadataCandidateLimit   = 30      # name-based candidates checked per unresolved label (one cheap label read each)


function Get-DmfProp {
    <#
    .SYNOPSIS  StrictMode-safe property read with a default.
    #>
    param([Parameter(Mandatory)][AllowNull()]$Object, [Parameter(Mandatory)][string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    return $p.Value
}


# =============================================================================
#  entity-map.json
# =============================================================================

function New-DmfEntityMapEntry {
    param([Parameter(Mandatory)][string]$EntityName)
    [pscustomobject]@{
        entityName            = $EntityName
        targetEntity          = $null
        publicEntityName      = $null
        publicCollectionName  = $null
        dataServiceEnabled    = $null
        dataManagementEnabled = $null
        entityCategory        = $null
        companySpecific       = $null
        keyFields             = @()
        source                = $null
        resolvedFrom          = $null
        resolvedAt            = $null
    }
}


function ConvertTo-DmfEntityMapEntry {
    param([Parameter(Mandatory)][string]$EntityName, [Parameter(Mandatory)]$Raw)
    $e = New-DmfEntityMapEntry -EntityName $EntityName
    foreach ($name in 'targetEntity', 'publicEntityName', 'publicCollectionName', 'entityCategory', 'source', 'resolvedFrom', 'resolvedAt') {
        $v = Get-DmfProp $Raw $name
        if ($null -ne $v -and [string]$v -ne '') { $e.$name = [string]$v }
    }
    foreach ($name in 'dataServiceEnabled', 'dataManagementEnabled', 'companySpecific') {
        $v = Get-DmfProp $Raw $name
        if ($null -ne $v) { $e.$name = [bool]$v }
    }
    $k = Get-DmfProp $Raw 'keyFields'
    if ($null -ne $k) { $e.keyFields = @($k | ForEach-Object { [string]$_ } | Where-Object { $_ }) }
    return $e
}


function Get-DmfEntityMap {
    <#
    .SYNOPSIS
        Loads resources/entity-map.json (or returns an empty map when absent).
    .OUTPUTS
        [pscustomobject]  Path, SchemaVersion, Entities (case-insensitive hashtable: label -> entry)
    #>
    param([Parameter(Mandatory)][string]$Path)

    $entities = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
    $map = [pscustomobject]@{ PSTypeName = 'Dmf.EntityMap'; Path = $Path; SchemaVersion = $Script:DmfEntityMapSchemaVersion; Entities = $entities }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $map }

    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    $rawEntities = Get-DmfProp $raw 'entities'
    if ($null -ne $rawEntities) {
        foreach ($p in $rawEntities.PSObject.Properties) {
            $entities[$p.Name] = ConvertTo-DmfEntityMapEntry -EntityName $p.Name -Raw $p.Value
        }
    }
    return $map
}


function Get-DmfEntityMapEntry {
    param([Parameter(Mandatory)]$Map, [Parameter(Mandatory)][string]$EntityName)
    if ($Map.Entities.ContainsKey($EntityName)) { return $Map.Entities[$EntityName] }
    return $null
}


function Find-DmfEntityMapByTarget {
    param([Parameter(Mandatory)]$Map, [Parameter(Mandatory)][string]$TargetEntity)
    foreach ($e in $Map.Entities.Values) {
        if ([string]::Equals([string]$e.targetEntity, $TargetEntity, [System.StringComparison]::OrdinalIgnoreCase)) { return $e }
    }
    return $null
}


function Test-DmfEntityMapResolved {
    <#
    .SYNOPSIS  True when the map can answer the OData question for a label without the network.
    #>
    param([Parameter(Mandatory)]$Map, [Parameter(Mandatory)][string]$EntityName)
    $e = Get-DmfEntityMapEntry -Map $Map -EntityName $EntityName
    if ($null -eq $e) { return $false }
    if ($e.dataServiceEnabled -eq $false) { return $true }                       # known not-public
    return (-not [string]::IsNullOrEmpty($e.publicCollectionName))
}


function Set-DmfEntityMapEntry {
    <#
    .SYNOPSIS
        Creates or updates a map entry.  A 'manual' entry only has its empty
        fields filled unless -Force is given.
    #>
    param(
        [Parameter(Mandatory)]$Map,
        [Parameter(Mandatory)][string]$EntityName,
        [Parameter(Mandatory)][hashtable]$Values,
        [switch]$Force
    )
    $e = Get-DmfEntityMapEntry -Map $Map -EntityName $EntityName
    if ($null -eq $e) { $e = New-DmfEntityMapEntry -EntityName $EntityName; $Map.Entities[$EntityName] = $e }
    $protect = ($e.source -eq 'manual' -and -not $Force)

    foreach ($k in $Values.Keys) {
        if ($null -eq $e.PSObject.Properties[$k]) { continue }
        $current = $e.$k
        $isEmpty = ($null -eq $current) -or (($current -is [string]) -and $current -eq '') -or (($current -is [array]) -and $current.Count -eq 0)
        if ($protect -and -not $isEmpty -and $k -ne 'resolvedAt') { continue }
        if ($protect -and $k -eq 'source') { continue }
        $e.$k = $Values[$k]
    }
    if (-not $Values.ContainsKey('resolvedAt')) {
        $e.resolvedAt = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    }
    return $e
}


function ConvertTo-DmfEntityMapDocument {
    param([Parameter(Mandatory)]$Map)
    $entities = [ordered]@{}
    foreach ($name in ($Map.Entities.Keys | Sort-Object)) {
        $e = $Map.Entities[$name]
        $entities[$name] = [ordered]@{
            targetEntity          = $e.targetEntity
            publicEntityName      = $e.publicEntityName
            publicCollectionName  = $e.publicCollectionName
            dataServiceEnabled    = $e.dataServiceEnabled
            dataManagementEnabled = $e.dataManagementEnabled
            entityCategory        = $e.entityCategory
            companySpecific       = $e.companySpecific
            keyFields             = @($e.keyFields)
            source                = $e.source
            resolvedFrom          = $e.resolvedFrom
            resolvedAt            = $e.resolvedAt
        }
    }
    return [ordered]@{ schemaVersion = $Script:DmfEntityMapSchemaVersion; entities = $entities }
}


function Save-DmfEntityMap {
    <#
    .SYNOPSIS
        Merges the in-memory map into the file on disk and writes it atomically.
    .DESCRIPTION
        Entries already on disk are re-read first so two runs cannot clobber
        each other's additions.  In-memory entries win, except that a 'manual'
        entry on disk is kept unless the in-memory entry is also 'manual'.
    #>
    param([Parameter(Mandatory)]$Map, [string]$Path)

    if (-not $Path) { $Path = $Map.Path }
    $disk = Get-DmfEntityMap -Path $Path
    foreach ($name in @($Map.Entities.Keys)) {
        $mem = $Map.Entities[$name]
        $old = $null
        if ($disk.Entities.ContainsKey($name)) { $old = $disk.Entities[$name] }
        if ($null -ne $old -and $old.source -eq 'manual' -and $mem.source -ne 'manual') { continue }
        $disk.Entities[$name] = $mem
    }

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir -PathType Container)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $json = (ConvertTo-DmfEntityMapDocument -Map $disk) | ConvertTo-Json -Depth 6
    $tmp  = "$Path.tmp"
    [System.IO.File]::WriteAllText($tmp, $json, (New-Object System.Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $tmp -Destination $Path -Force

    # keep the caller's map in sync with what was written
    foreach ($name in @($disk.Entities.Keys)) { $Map.Entities[$name] = $disk.Entities[$name] }
    return $Path
}


function Update-DmfEntityMapFromManifests {
    <#
    .SYNOPSIS
        Harvests EntityName -> TargetEntity pairs from every Manifest.xml under
        the given paths into the map (source 'manifest').  No network.
    .OUTPUTS
        [pscustomobject]  Files, Lines, Added, Updated, AlreadyKnown
    #>
    param(
        [Parameter(Mandatory)][string[]]$Path,
        [Parameter(Mandatory)]$Map
    )
    $files = 0; $lines = 0; $added = 0; $updated = 0; $known = 0
    foreach ($root in $Path) {
        if (-not (Test-Path -LiteralPath $root)) { Write-Warn "Seed path not found: '$root'"; continue }
        $manifests = @(Get-ChildItem -LiteralPath $root -Filter 'Manifest.xml' -File -Recurse -ErrorAction SilentlyContinue)
        foreach ($mf in $manifests) {
            $m = $null
            try { $m = Read-DmfManifest -Path $mf.FullName } catch { Write-Warn "Skipping '$($mf.FullName)': $($_.Exception.Message)"; continue }
            $files++
            foreach ($l in $m.Lines) {
                $lines++
                if ([string]::IsNullOrWhiteSpace($l.TargetEntity)) { continue }
                $e = Get-DmfEntityMapEntry -Map $Map -EntityName $l.EntityName
                if ($null -eq $e) {
                    [void](Set-DmfEntityMapEntry -Map $Map -EntityName $l.EntityName -Values @{ targetEntity = $l.TargetEntity; source = 'manifest' })
                    $added++
                } elseif ([string]::IsNullOrWhiteSpace($e.targetEntity)) {
                    $vals = @{ targetEntity = $l.TargetEntity }
                    if (-not $e.source) { $vals['source'] = 'manifest' }
                    [void](Set-DmfEntityMapEntry -Map $Map -EntityName $l.EntityName -Values $vals)
                    $updated++
                } else { $known++ }
            }
        }
    }
    return [pscustomobject]@{ Files = $files; Lines = $lines; Added = $added; Updated = $updated; AlreadyKnown = $known }
}


# =============================================================================
#  Metadata service wrappers
# =============================================================================

function Get-DmfMetadataDataEntities {
    <#
    .SYNOPSIS  All /Metadata/DataEntities rows, fetched once per run per environment.
    #>
    param([Parameter(Mandatory)]$Session, [switch]$Refresh)

    $key = $Session.BaseUrl
    if (-not $Refresh -and $Script:DmfMetadataEntityCache.ContainsKey($key)) { return $Script:DmfMetadataEntityCache[$key] }

    $headers = Get-DmfAuthHeaders -Session $Session
    $select  = 'Name,PublicEntityName,PublicCollectionName,LabelId,DataServiceEnabled,DataManagementEnabled,EntityCategory,IsReadOnly'
    $rows = $null
    try {
        $uri  = New-DmfODataUri -BaseUrl $Session.BaseUrl -ServicePath 'Metadata' -Collection 'DataEntities' -Select $select
        $rows = @(Get-DmfODataAll -Uri $uri -Operation 'Metadata DataEntities' -Headers $headers)
    } catch {
        Write-Detail "Metadata DataEntities with `$select failed ($($_.Exception.Message)); retrying without `$select."
        $uri  = New-DmfODataUri -BaseUrl $Session.BaseUrl -ServicePath 'Metadata' -Collection 'DataEntities'
        $rows = @(Get-DmfODataAll -Uri $uri -Operation 'Metadata DataEntities' -Headers $headers)
    }
    $Script:DmfMetadataEntityCache[$key] = $rows
    Write-Detail "Metadata service: $($rows.Count) data entities loaded."
    return $rows
}


function Get-DmfMetadataDataEntity {
    <#
    .SYNOPSIS
        One DataEntities row by AOT name: the cached full list when already
        loaded, otherwise $filter, then the key segment, then a full-list scan.
    #>
    param([Parameter(Mandatory)]$Session, [Parameter(Mandatory)][string]$Name)

    $scan = {
        foreach ($r in (Get-DmfMetadataDataEntities -Session $Session)) {
            if ([string]::Equals([string](Get-DmfProp $r 'Name' ''), $Name, [System.StringComparison]::OrdinalIgnoreCase)) { return $r }
        }
        return $null
    }
    if ($Script:DmfMetadataEntityCache.ContainsKey($Session.BaseUrl)) { return (& $scan) }

    $headers = Get-DmfAuthHeaders -Session $Session
    try {
        $uri  = New-DmfODataUri -BaseUrl $Session.BaseUrl -ServicePath 'Metadata' -Collection 'DataEntities' -Filter "Name eq '$(ConvertTo-DmfODataLiteral $Name)'"
        $rows = @(Get-DmfODataAll -Uri $uri -Operation "Metadata DataEntities $Name" -Headers $headers)
        if ($rows.Count -gt 0) { return $rows[0] }
        return $null
    } catch {
        Write-Detail "Metadata DataEntities `$filter failed ($($_.Exception.Message)); trying the key segment."
    }
    try {
        $uri  = "$($Session.BaseUrl)/Metadata/DataEntities('$([System.Uri]::EscapeDataString($Name))')"
        $rows = @(Get-DmfODataAll -Uri $uri -Operation "Metadata DataEntities('$Name')" -Headers $headers)
        if ($rows.Count -gt 0) { return $rows[0] }
    } catch {
        if ($_.Exception.Message -match 'HTTP 404') { return $null }
        Write-Detail "Metadata DataEntities key segment failed ($($_.Exception.Message)); scanning the full list."
    }
    return (& $scan)
}


function Get-DmfMetadataLabel {
    <#
    .SYNOPSIS
        Text of one label via the key segment /Metadata/Labels(Id='...',Language='...').
    .DESCRIPTION
        The Metadata service rejects $filter on Labels (HTTP 501 "The label id
        and language must be provided as part of the key segment"), so labels
        are read one at a time and cached for the run.  Returns $null when the
        label does not exist.
    #>
    param([Parameter(Mandatory)]$Session, [Parameter(Mandatory)][string]$LabelId, [string]$Language = 'en-US')

    $cacheKey = "$Language|$LabelId"
    if ($Script:DmfMetadataLabelTextCache.ContainsKey($cacheKey)) { return $Script:DmfMetadataLabelTextCache[$cacheKey] }

    $headers = Get-DmfAuthHeaders -Session $Session
    $uri = "$($Session.BaseUrl)/Metadata/Labels(Id='$([System.Uri]::EscapeDataString((ConvertTo-DmfODataLiteral $LabelId)))',Language='$([System.Uri]::EscapeDataString((ConvertTo-DmfODataLiteral $Language)))')"
    $text = $null
    try {
        $rows = @(Get-DmfODataAll -Uri $uri -Operation "Metadata Label $LabelId" -Headers $headers)
        if ($rows.Count -gt 0) {
            $v = Get-DmfProp $rows[0] 'Value'
            if ($null -ne $v) { $text = [string]$v }
        }
    } catch {
        if ($_.Exception.Message -notmatch 'HTTP 404') { Write-Detail "Metadata Label $LabelId failed: $($_.Exception.Message)" }
    }
    $Script:DmfMetadataLabelTextCache[$cacheKey] = $text
    return $text
}


function Get-DmfMetadataDataEntitiesByLabelId {
    param([Parameter(Mandatory)]$Session, [Parameter(Mandatory)][string[]]$LabelIds)

    $headers = Get-DmfAuthHeaders -Session $Session
    try {
        $filter = ($LabelIds | ForEach-Object { "LabelId eq '$(ConvertTo-DmfODataLiteral $_)'" }) -join ' or '
        $uri    = New-DmfODataUri -BaseUrl $Session.BaseUrl -ServicePath 'Metadata' -Collection 'DataEntities' -Filter $filter
        return @(Get-DmfODataAll -Uri $uri -Operation 'Metadata DataEntities by LabelId' -Headers $headers)
    } catch {
        Write-Detail "Metadata DataEntities LabelId `$filter failed ($($_.Exception.Message)); scanning the full list."
        $set = [System.Collections.Generic.HashSet[string]]::new([string[]]$LabelIds, [System.StringComparer]::OrdinalIgnoreCase)
        return @((Get-DmfMetadataDataEntities -Session $Session) | Where-Object { $set.Contains([string](Get-DmfProp $_ 'LabelId' '')) })
    }
}


function Get-DmfMetadataLabels {
    <#
    .SYNOPSIS
        Resolves label ids to text.  Returns a hashtable Id -> Value.
    .DESCRIPTION
        Tries one or-chained $filter batch first (cheap when a service version
        supports it); as soon as the service rejects that form it switches to
        key-segment reads (Get-DmfMetadataLabel) for the rest of the run.
    #>
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$LabelIds,
        [string]$Language = 'en-US'
    )
    $result = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
    $ids    = @($LabelIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($ids.Count -eq 0) { return $result }

    $pending = [System.Collections.Generic.List[string]]::new()
    foreach ($id in $ids) {
        $k = "$Language|$id"
        if ($Script:DmfMetadataLabelTextCache.ContainsKey($k)) {
            if ($null -ne $Script:DmfMetadataLabelTextCache[$k]) { $result[$id] = $Script:DmfMetadataLabelTextCache[$k] }
        } else { $pending.Add($id) }
    }
    if ($pending.Count -eq 0) { return $result }

    if ($Script:DmfMetadataLabelFilterOk -ne $false) {
        $headers = Get-DmfAuthHeaders -Session $Session
        $lang    = ConvertTo-DmfODataLiteral $Language
        try {
            for ($i = 0; $i -lt $pending.Count; $i += $Script:DmfLabelBatchSize) {
                $batch = @($pending[$i..([Math]::Min($i + $Script:DmfLabelBatchSize, $pending.Count) - 1)])
                $chain = ($batch | ForEach-Object { "Id eq '$(ConvertTo-DmfODataLiteral $_)'" }) -join ' or '
                $uri   = New-DmfODataUri -BaseUrl $Session.BaseUrl -ServicePath 'Metadata' -Collection 'Labels' -Filter "($chain) and Language eq '$lang'"
                $rows  = @(Get-DmfODataAll -Uri $uri -Operation "Metadata Labels batch $([int]($i / $Script:DmfLabelBatchSize) + 1)" -Headers $headers)
                foreach ($id in $batch) { $Script:DmfMetadataLabelTextCache["$Language|$id"] = $null }
                foreach ($r in $rows) {
                    $id = [string](Get-DmfProp $r 'Id' '')
                    if (-not $id) { continue }
                    $text = [string](Get-DmfProp $r 'Value' '')
                    $result[$id] = $text
                    $Script:DmfMetadataLabelTextCache["$Language|$id"] = $text
                }
                $Script:DmfMetadataLabelFilterOk = $true
            }
            return $result
        } catch {
            $Script:DmfMetadataLabelFilterOk = $false
            Write-Detail "Metadata Labels `$filter not supported ($($_.Exception.Message)); reading labels by key segment."
        }
    }

    foreach ($id in $pending) {
        $text = Get-DmfMetadataLabel -Session $Session -LabelId $id -Language $Language
        if ($null -ne $text) { $result[$id] = $text }
    }
    return $result
}


function Find-DmfMetadataLabelIds {
    <#
    .SYNOPSIS  Reverse lookup: label text -> label ids.  Throws when the Value filter is unsupported.
    #>
    param([Parameter(Mandatory)]$Session, [Parameter(Mandatory)][string]$Value, [string]$Language = 'en-US')
    $headers = Get-DmfAuthHeaders -Session $Session
    $uri  = New-DmfODataUri -BaseUrl $Session.BaseUrl -ServicePath 'Metadata' -Collection 'Labels' -Filter "Value eq '$(ConvertTo-DmfODataLiteral $Value)' and Language eq '$(ConvertTo-DmfODataLiteral $Language)'"
    $rows = @(Get-DmfODataAll -Uri $uri -Operation "Metadata Labels by value" -Headers $headers)
    return @($rows | ForEach-Object { [string](Get-DmfProp $_ 'Id' '') } | Where-Object { $_ } | Select-Object -Unique)
}


function Get-DmfMetadataPublicEntity {
    <#
    .SYNOPSIS
        One /Metadata/PublicEntities row (with Properties[]) by public entity
        name: $filter, then the key segment, then a one-time full-list scan.
    #>
    param([Parameter(Mandatory)]$Session, [Parameter(Mandatory)][string]$Name)

    $headers = Get-DmfAuthHeaders -Session $Session
    if (-not $Script:DmfMetadataPublicCache.ContainsKey($Session.BaseUrl)) {
        try {
            $uri  = New-DmfODataUri -BaseUrl $Session.BaseUrl -ServicePath 'Metadata' -Collection 'PublicEntities' -Filter "Name eq '$(ConvertTo-DmfODataLiteral $Name)'"
            $rows = @(Get-DmfODataAll -Uri $uri -Operation "Metadata PublicEntities $Name" -Headers $headers)
            if ($rows.Count -gt 0) { return $rows[0] }
            return $null
        } catch {
            Write-Detail "Metadata PublicEntities `$filter failed ($($_.Exception.Message)); trying the key segment."
        }
        try {
            $uri  = "$($Session.BaseUrl)/Metadata/PublicEntities('$([System.Uri]::EscapeDataString($Name))')"
            $rows = @(Get-DmfODataAll -Uri $uri -Operation "Metadata PublicEntities('$Name')" -Headers $headers)
            if ($rows.Count -gt 0) { return $rows[0] }
        } catch {
            if ($_.Exception.Message -match 'HTTP 404') { return $null }
            Write-Detail "Metadata PublicEntities key segment failed ($($_.Exception.Message)); loading the full list once."
        }
        $uri = New-DmfODataUri -BaseUrl $Session.BaseUrl -ServicePath 'Metadata' -Collection 'PublicEntities'
        $Script:DmfMetadataPublicCache[$Session.BaseUrl] = @(Get-DmfODataAll -Uri $uri -Operation 'Metadata PublicEntities (all)' -Headers $headers)
        Write-Detail "Metadata service: $($Script:DmfMetadataPublicCache[$Session.BaseUrl].Count) public entities loaded."
    }
    foreach ($r in $Script:DmfMetadataPublicCache[$Session.BaseUrl]) {
        if ([string]::Equals([string](Get-DmfProp $r 'Name' ''), $Name, [System.StringComparison]::OrdinalIgnoreCase)) { return $r }
    }
    return $null
}


# =============================================================================
#  Resolution
# =============================================================================

function Get-DmfNameTokens {
    <#
    .SYNOPSIS
        Lower-case word tokens of a label or an AOT entity name, with crude
        plural stripping, for candidate matching.  'Customer groups' and
        'CustCustomerGroupEntity' both yield tokens containing 'customer' and 'group'.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)
    $stop = @('and', 'of', 'the', 'for', 'to', 'in', 'a', 'an', 'entity', 'v2', 'v3', 'v4', 'v5')
    $spaced = [regex]::Replace($Text, '([a-z0-9])([A-Z])', '$1 $2')          # split camel case
    $spaced = [regex]::Replace($spaced, '([A-Z]+)([A-Z][a-z])', '$1 $2')      # ABCDef -> ABC Def
    $tokens = [System.Collections.Generic.List[string]]::new()
    foreach ($w in ($spaced -split '[^A-Za-z0-9]+')) {
        $t = $w.ToLowerInvariant()
        if ($t.Length -lt 2 -or $stop -contains $t) { continue }
        if ($t.Length -gt 3 -and $t.EndsWith('ies')) { $t = $t.Substring(0, $t.Length - 3) + 'y' }
        elseif ($t.Length -gt 3 -and $t.EndsWith('es') -and $t -notmatch '(ss|sh|ch|x|z)es$') { $t = $t.Substring(0, $t.Length - 1) }
        if ($t.Length -gt 3 -and $t.EndsWith('s') -and -not $t.EndsWith('ss')) { $t = $t.Substring(0, $t.Length - 1) }
        $tokens.Add($t)
    }
    return $tokens.ToArray()
}


function Find-DmfEntityCandidatesByName {
    <#
    .SYNOPSIS
        Data entities whose AOT name shares words with a DMF label, best first.
    .DESCRIPTION
        Replaces a full label index (which would need one key-segment read per
        entity, thousands of calls): only these few candidates have their label
        text read and compared exactly.  Score = share of the label's tokens
        found in the entity name; ties prefer DataManagementEnabled entities
        and shorter names.
    #>
    param(
        [Parameter(Mandatory)][string]$Label,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Entities,
        [int]$Limit = $Script:DmfMetadataCandidateLimit
    )
    $labelTokens = @(Get-DmfNameTokens -Text $Label | Select-Object -Unique)
    if ($labelTokens.Count -eq 0) { return @() }

    $scored = foreach ($e in $Entities) {
        $name = [string](Get-DmfProp $e 'Name' '')
        if (-not $name) { continue }
        # Words from the AOT name plus the public entity / collection names:
        # 'Global address book V2' matches DirPartyEntity only through its
        # public name, GlobalAddressBook.
        $nameTokens = @(@(Get-DmfNameTokens -Text $name) +
                        @(Get-DmfNameTokens -Text ([string](Get-DmfProp $e 'PublicEntityName' ''))) +
                        @(Get-DmfNameTokens -Text ([string](Get-DmfProp $e 'PublicCollectionName' ''))) | Select-Object -Unique)
        if ($nameTokens.Count -eq 0) { continue }
        $hits = 0
        foreach ($lt in $labelTokens) {
            foreach ($nt in $nameTokens) { if ($nt -eq $lt -or $nt.StartsWith($lt) -or $lt.StartsWith($nt)) { $hits++; break } }
        }
        $score = $hits / $labelTokens.Count
        if ($score -lt 0.5) { continue }
        [pscustomobject]@{ Entity = $e; Score = $score; Dm = [bool](Get-DmfProp $e 'DataManagementEnabled' $false); Len = $name.Length }
    }
    return @($scored | Sort-Object -Property @{ e = 'Score'; d = $true }, @{ e = 'Dm'; d = $true }, @{ e = 'Len'; d = $false } |
        Select-Object -First $Limit | ForEach-Object { $_.Entity })
}


function Find-DmfEntitiesByLabelText {
    <#
    .SYNOPSIS
        Data entities whose label text equals the given DMF label, found by
        checking name-based candidates with key-segment label reads.
    #>
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$Label,
        [string]$Language = 'en-US'
    )
    $all   = @(Get-DmfMetadataDataEntities -Session $Session)
    $cands = @(Find-DmfEntityCandidatesByName -Label $Label -Entities $all)
    if ($cands.Count -eq 0) { return @() }
    $ids   = @($cands | ForEach-Object { [string](Get-DmfProp $_ 'LabelId' '') } | Where-Object { $_ } | Select-Object -Unique)
    $texts = Get-DmfMetadataLabels -Session $Session -LabelIds $ids -Language $Language
    $want  = $Label.Trim()
    $hits  = @($cands | Where-Object {
        $id = [string](Get-DmfProp $_ 'LabelId' '')
        $id -and $texts.ContainsKey($id) -and [string]::Equals(([string]$texts[$id]).Trim(), $want, [System.StringComparison]::OrdinalIgnoreCase)
    })
    if ($hits.Count -eq 0) {
        # Diagnostics for tuning: what was checked and what its label said.
        $tried = @($cands | Select-Object -First 6 | ForEach-Object {
            $id = [string](Get-DmfProp $_ 'LabelId' '')
            "$(Get-DmfProp $_ 'Name' '')='$(if ($id -and $texts.ContainsKey($id)) { $texts[$id] } else { '?' })'"
        })
        Write-Detail "No label match for '$Label' among $($cands.Count) name candidates: $($tried -join ', ')"
    }
    return $hits
}


function Resolve-DmfEntity {
    <#
    .SYNOPSIS
        Resolves template lines to OData collections, keys and company flags.

    .PARAMETER Lines
        Objects with EntityName and (optionally) TargetEntity -- manifest lines
        from Read-DmfManifest or converted template lines.

    .PARAMETER Session
        Dmf.Session; required unless -Offline.

    .PARAMETER Map
        Entity map from Get-DmfEntityMap.  Updated in place; call
        Save-DmfEntityMap afterwards to persist.

    .PARAMETER Refresh
        Ignore cached resolutions (manual entries are still honoured).

    .PARAMETER Offline
        Cache only; misses are reported as Unresolved.

    .OUTPUTS
        One object per line: EntityName, TargetEntity, PublicEntityName,
        Collection, KeyFields, CompanySpecific, DataServiceEnabled,
        Status (Resolved | NotPublic | Unresolved | Unresolved-Ambiguous),
        Reason, Source (cache | metadata | manual | none)
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Lines,
        $Session,
        [Parameter(Mandatory)]$Map,
        [switch]$Refresh,
        [switch]$Offline,
        [string]$Language = 'en-US',
        [string]$ResolvedFrom
    )

    if (-not $Offline -and $null -eq $Session) { throw 'Resolve-DmfEntity needs -Session unless -Offline is given.' }
    if (-not $ResolvedFrom -and $null -ne $Session) { $ResolvedFrom = $Session.EnvironmentName }

    $valueFilterSupported = $true

    foreach ($line in $Lines) {
        $name   = [string](Get-DmfProp $line 'EntityName' '')
        $target = Get-DmfProp $line 'TargetEntity'
        if ([string]::IsNullOrWhiteSpace($name)) { continue }

        $res = [pscustomobject]@{
            EntityName         = $name
            TargetEntity       = $target
            PublicEntityName   = $null
            Collection         = $null
            KeyFields          = @()
            CompanySpecific    = $false
            DataServiceEnabled = $null
            Status             = 'Unresolved'
            Reason             = ''
            Source             = 'none'
        }

        $entry = Get-DmfEntityMapEntry -Map $Map -EntityName $name
        $useCache = ($null -ne $entry) -and ((-not $Refresh) -or $entry.source -eq 'manual') -and (Test-DmfEntityMapResolved -Map $Map -EntityName $name)
        if ($useCache) {
            $res.TargetEntity       = $entry.targetEntity
            $res.PublicEntityName   = $entry.publicEntityName
            $res.Collection         = $entry.publicCollectionName
            $res.KeyFields          = @($entry.keyFields)
            $res.CompanySpecific    = [bool]$entry.companySpecific
            $res.DataServiceEnabled = $entry.dataServiceEnabled
            $res.Source             = if ($entry.source -eq 'manual') { 'manual' } else { 'cache' }
            if ($entry.dataServiceEnabled -eq $false -or -not $entry.publicCollectionName) {
                $res.Status = 'NotPublic'; $res.Reason = 'DataServiceEnabled=false'
            } else {
                $res.Status = 'Resolved'
            }
            Write-Output $res
            continue
        }

        if ($Offline) {
            if ($null -ne $entry -and $entry.targetEntity) { $res.TargetEntity = $entry.targetEntity }
            $res.Reason = 'not in entity map (offline)'
            Write-Output $res
            continue
        }

        try {
            if (-not $target -and $null -ne $entry -and $entry.targetEntity) { $target = $entry.targetEntity }
            $de = $null

            if ($target) {
                $de = Get-DmfMetadataDataEntity -Session $Session -Name $target
                if ($null -eq $de) { $res.Reason = "TargetEntity '$target' not found in the Metadata service" }
            }
            else {
                $cands = @()
                $ids   = $null
                # 1. Reverse lookup by label text -- only some service versions
                #    accept $filter on Labels; the first refusal disables it for
                #    the rest of the run.
                if ($valueFilterSupported -and $Script:DmfMetadataLabelFilterOk -ne $false) {
                    try { $ids = @(Find-DmfMetadataLabelIds -Session $Session -Value $name -Language $Language) }
                    catch {
                        $valueFilterSupported = $false
                        $Script:DmfMetadataLabelFilterOk = $false
                        Write-Detail "Labels Value filter not supported ($($_.Exception.Message)); matching by entity name and key-segment label reads instead."
                    }
                } else { $valueFilterSupported = $false }
                if ($valueFilterSupported) {
                    if ($ids.Count -gt 0) { $cands = @(Get-DmfMetadataDataEntitiesByLabelId -Session $Session -LabelIds $ids) }
                } else {
                    # 2. Name-based candidates confirmed by exact label text.
                    $cands = @(Find-DmfEntitiesByLabelText -Session $Session -Label $name -Language $Language)
                }

                $dm = @($cands | Where-Object { [bool](Get-DmfProp $_ 'DataManagementEnabled' $false) })
                # @() around the if: a one-element result would otherwise unroll
                # to a scalar, which has no .Count under StrictMode on PS 5.1.
                $pool = @(if ($dm.Count -gt 0) { $dm } else { $cands })
                if ($pool.Count -eq 1) { $de = $pool[0] }
                elseif ($pool.Count -gt 1) {
                    $res.Status = 'Unresolved-Ambiguous'
                    $res.Reason = 'label matches several entities: ' + (($pool | ForEach-Object { Get-DmfProp $_ 'Name' '' }) -join ', ')
                }
                else { $res.Reason = 'no data entity with this label found (seed TargetEntity from an exported Manifest.xml with -SeedFromPath, or add a manual entity-map.json entry)' }
            }

            if ($null -ne $de) {
                $aot   = [string](Get-DmfProp $de 'Name' '')
                $dse   = [bool](Get-DmfProp $de 'DataServiceEnabled' $false)
                $dme   = [bool](Get-DmfProp $de 'DataManagementEnabled' $false)
                $pen   = [string](Get-DmfProp $de 'PublicEntityName' '')
                $pcn   = [string](Get-DmfProp $de 'PublicCollectionName' '')
                $cat   = [string](Get-DmfProp $de 'EntityCategory' '')
                $keys  = @()
                $company = $false

                if ($dse -and $pen) {
                    $pe = Get-DmfMetadataPublicEntity -Session $Session -Name $pen
                    if ($null -ne $pe) {
                        $props   = @(Get-DmfProp $pe 'Properties' @())
                        $keys    = @($props | Where-Object { [bool](Get-DmfProp $_ 'IsKey' $false) } | ForEach-Object { [string](Get-DmfProp $_ 'Name' '') } | Where-Object { $_ })
                        $company = (@($props | Where-Object { [string](Get-DmfProp $_ 'Name' '') -eq 'dataAreaId' }).Count -gt 0)
                        if (-not $pcn) { $pcn = [string](Get-DmfProp $pe 'EntitySetName' '') }
                    }
                }

                [void](Set-DmfEntityMapEntry -Map $Map -EntityName $name -Values @{
                    targetEntity          = $aot
                    publicEntityName      = $(if ($pen) { $pen } else { $null })
                    publicCollectionName  = $(if ($pcn) { $pcn } else { $null })
                    dataServiceEnabled    = $dse
                    dataManagementEnabled = $dme
                    entityCategory        = $(if ($cat) { $cat } else { $null })
                    companySpecific       = $company
                    keyFields             = $keys
                    source                = 'metadata'
                    resolvedFrom          = $ResolvedFrom
                })

                $res.TargetEntity       = $aot
                $res.PublicEntityName   = $pen
                $res.Collection         = $pcn
                $res.KeyFields          = $keys
                $res.CompanySpecific    = $company
                $res.DataServiceEnabled = $dse
                $res.Source             = 'metadata'
                if ($dse -and $pcn) { $res.Status = 'Resolved'; $res.Reason = '' }
                else { $res.Status = 'NotPublic'; $res.Reason = 'DataServiceEnabled=false' }
            }
        }
        catch {
            $res.Status = 'Unresolved'
            $res.Reason = "Metadata service error: $($_.Exception.Message)"
        }

        Write-Output $res
    }
}
