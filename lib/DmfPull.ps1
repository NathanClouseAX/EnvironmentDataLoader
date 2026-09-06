<#
.SYNOPSIS
    OData snapshot storage for the data pull: data/<env>/<LE>/<Entity>.json
    and the _pull.json run index.

.DESCRIPTION
    Dot-source this file (after DmfOutput.ps1, DmfRequest.ps1, DmfAuth.ps1,
    DmfOData.ps1, DmfTemplate.ps1) to import:

        Get-DmfPullFolder          -- data/<env>/<LE>/ path (optionally created)
        Get-DmfSnapshotFileName    -- '<Entity label>.json' with invalid chars removed
        Invoke-DmfEntityPull       -- read one entity's rows via OData (paged)
        Write-DmfEntitySnapshot    -- write the entity file (sorted, atomic)
        Read-DmfEntitySnapshot     -- read an entity file back (used by the diff)
        Read-DmfPullIndex / Write-DmfPullIndex / Set-DmfPullIndexEntry

    Entity file (<Entity label>.json): an envelope -- schemaVersion, entity,
    targetEntity, collection, environment, environmentUrl, legalEntity,
    companySpecific, keyFields, pulledAt, elapsedSeconds, recordCount,
    truncated, fields (union of property names in first-seen order) -- and
    'records', the OData rows verbatim, sorted by key.

    Run index (_pull.json): environment, environmentUrl, legalEntity, lastRun
    (startedAt, finishedAt, templates, tool) and one entry per entity with
    status (Pulled | Truncated | NotPublic | Unresolved | Failed), reason,
    file, recordCount, pulledAt and the templates that pulled it.  Merged on
    every run, so a folder can accumulate several templates; a failed pull
    keeps the previous good file.
#>

function Get-DmfSnapshotFileName {
    param([Parameter(Mandatory)][string]$EntityName)
    return (ConvertTo-DmfSafeFileName -Name $EntityName) + '.json'
}


function Get-DmfPullFolder {
    param(
        [Parameter(Mandatory)][string]$DataPath,
        [Parameter(Mandatory)][string]$EnvironmentName,
        [Parameter(Mandatory)][string]$LegalEntityId,
        [switch]$Create
    )
    $folder = Join-Path (Join-Path $DataPath $EnvironmentName) $LegalEntityId.ToUpperInvariant()
    if ($Create -and -not (Test-Path -LiteralPath $folder -PathType Container)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    return $folder
}


function Write-DmfJsonAtomic {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Json)
    $tmp = "$Path.tmp"
    [System.IO.File]::WriteAllText($tmp, $Json, (New-Object System.Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $tmp -Destination $Path -Force
}


function Invoke-DmfEntityPull {
    <#
    .SYNOPSIS
        Reads every row of one entity for one legal entity.

    .PARAMETER Resolution
        Object from Resolve-DmfEntity (Collection, CompanySpecific, EntityName).

    .OUTPUTS
        [pscustomobject]  Records (array), Truncated, Uri, Pages
    #>
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)]$Resolution,
        [Parameter(Mandatory)][string]$LegalEntityId,
        [int]$MaxRecords = 0,
        [scriptblock]$OnPage
    )

    $uriArgs = @{ BaseUrl = $Session.BaseUrl; Collection = $Resolution.Collection }
    if ($Resolution.CompanySpecific) {
        # Without cross-company F&O silently scopes to the caller's default
        # company; the filter alone would then make the folder name a lie.
        $uriArgs['CrossCompany'] = $true
        $uriArgs['Filter']       = "dataAreaId eq '$(ConvertTo-DmfODataLiteral $LegalEntityId)'"
    }
    $uri     = New-DmfODataUri @uriArgs
    $headers = Get-DmfAuthHeaders -Session $Session

    # The caller's callback is handed straight to Get-DmfODataAll.  Do not wrap
    # it in another scriptblock that mentions $OnPage: a plain scriptblock runs
    # in the callee's scope, where $OnPage is the callee's own parameter -- the
    # wrapper would call itself until the call stack overflows.
    $pullArgs = @{ Uri = $uri; Operation = "pull $($Resolution.EntityName)"; Headers = $headers; MaxItems = $MaxRecords }
    if ($null -ne $OnPage) { $pullArgs['OnPage'] = $OnPage }

    $records   = @(Get-DmfODataAll @pullArgs)
    $truncated = ($MaxRecords -gt 0 -and $records.Count -ge $MaxRecords)

    return [pscustomobject]@{ Records = $records; Truncated = $truncated; Uri = $uri }
}


function Get-DmfRecordKeyString {
    param([Parameter(Mandatory)]$Record, [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$KeyFields)
    $parts = foreach ($k in $KeyFields) {
        $p = $Record.PSObject.Properties[$k]
        if ($null -ne $p -and $null -ne $p.Value) { [string]$p.Value } else { '' }
    }
    return (@($parts) -join [char]0x1F)
}


function Write-DmfEntitySnapshot {
    <#
    .SYNOPSIS
        Writes <Entity>.json for one pulled entity (records verbatim, sorted by key).
    .OUTPUTS
        [string]  the file path written.
    #>
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)]$Resolution,
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$LegalEntityId,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Records,
        [double]$ElapsedSeconds = 0,
        [bool]$Truncated = $false
    )

    $keys = @($Resolution.KeyFields)
    # @() around the if: an empty or one-element result would otherwise
    # collapse to $null / a scalar, which has no .Count under StrictMode on PS 5.1.
    $sorted = @(if ($keys.Count -gt 0) {
        $Records | Sort-Object -Property @{ Expression = { Get-DmfRecordKeyString -Record $_ -KeyFields $keys } }
    } else { $Records })

    # Union of field names in first-seen order
    $fields = [System.Collections.Generic.List[string]]::new()
    $seen   = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    foreach ($r in $sorted) {
        foreach ($p in $r.PSObject.Properties) { if ($seen.Add($p.Name)) { $fields.Add($p.Name) } }
    }

    $envelope = [ordered]@{
        schemaVersion   = 1
        entity          = $Resolution.EntityName
        targetEntity    = $Resolution.TargetEntity
        collection      = $Resolution.Collection
        environment     = $Session.EnvironmentName
        environmentUrl  = $Session.BaseUrl
        legalEntity     = $LegalEntityId.ToUpperInvariant()
        companySpecific = [bool]$Resolution.CompanySpecific
        keyFields       = $keys
        pulledAt        = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        elapsedSeconds  = [Math]::Round($ElapsedSeconds, 1)
        recordCount     = $sorted.Count
        truncated       = $Truncated
        fields          = $fields.ToArray()
        records         = $sorted
    }

    $path = Join-Path $Folder (Get-DmfSnapshotFileName -EntityName $Resolution.EntityName)
    Write-DmfJsonAtomic -Path $path -Json ($envelope | ConvertTo-Json -Depth 10)
    return $path
}


function Read-DmfEntitySnapshot {
    <#
    .SYNOPSIS  Reads an entity snapshot; records / keyFields / fields are always arrays.
    #>
    param([Parameter(Mandatory)][string]$Path)
    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    $get = { param($n, $d) $p = $raw.PSObject.Properties[$n]; if ($null -ne $p -and $null -ne $p.Value) { $p.Value } else { $d } }
    # PowerShell 7 turns ISO-8601 strings into [DateTime] on parse; keep the
    # metadata timestamp textual so both hosts present it the same way.
    $pulledAt = & $get 'pulledAt' $null
    if ($pulledAt -is [DateTime]) { $pulledAt = $pulledAt.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
    return [pscustomobject]@{
        Path            = $Path
        Entity          = [string](& $get 'entity' ([System.IO.Path]::GetFileNameWithoutExtension($Path)))
        TargetEntity    = & $get 'targetEntity' $null
        Collection      = & $get 'collection' $null
        Environment     = & $get 'environment' $null
        EnvironmentUrl  = & $get 'environmentUrl' $null
        LegalEntity     = & $get 'legalEntity' $null
        CompanySpecific = [bool](& $get 'companySpecific' $false)
        KeyFields       = @(& $get 'keyFields' @())
        PulledAt        = $pulledAt
        RecordCount     = [int](& $get 'recordCount' 0)
        Truncated       = [bool](& $get 'truncated' $false)
        Fields          = @(& $get 'fields' @())
        Records         = @(& $get 'records' @())
    }
}


function Read-DmfPullIndex {
    <#
    .SYNOPSIS  Reads _pull.json into mutable hashtables (or returns a new, empty index).
    #>
    param([Parameter(Mandatory)][string]$Folder)
    $path  = Join-Path $Folder '_pull.json'
    $index = [ordered]@{
        schemaVersion  = 1
        environment    = $null
        environmentUrl = $null
        legalEntity    = $null
        lastRun        = $null
        entities       = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
    }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $index }
    try {
        $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($n in 'environment', 'environmentUrl', 'legalEntity') {
            $p = $raw.PSObject.Properties[$n]; if ($null -ne $p) { $index[$n] = $p.Value }
        }
        $lr = $raw.PSObject.Properties['lastRun']; if ($null -ne $lr) { $index.lastRun = $lr.Value }
        $ents = $raw.PSObject.Properties['entities']
        if ($null -ne $ents -and $null -ne $ents.Value) {
            foreach ($p in $ents.Value.PSObject.Properties) {
                $h = [ordered]@{}
                foreach ($q in $p.Value.PSObject.Properties) { $h[$q.Name] = $q.Value }
                $index.entities[$p.Name] = $h
            }
        }
    } catch {
        Write-Warn "_pull.json in '$Folder' could not be parsed ($($_.Exception.Message)); starting a new index."
    }
    return $index
}


function Set-DmfPullIndexEntry {
    param(
        [Parameter(Mandatory)]$Index,
        [Parameter(Mandatory)][string]$EntityName,
        [Parameter(Mandatory)][string]$Status,
        [string]$Reason,
        [string]$File,
        [int]$RecordCount = -1,
        [string[]]$Templates = @()
    )
    $entry = [ordered]@{ status = $Status }
    if ($Reason)           { $entry.reason = $Reason }
    if ($File)             { $entry.file = $File }
    if ($RecordCount -ge 0) { $entry.recordCount = $RecordCount }
    $entry.pulledAt  = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $entry.templates = @($Templates)
    # a failed pull keeps the previous file reference so the stale file is still findable
    if ($Status -eq 'Failed' -and $Index.entities.ContainsKey($EntityName)) {
        $old = $Index.entities[$EntityName]
        if ($old.Contains('file') -and -not $File) { $entry.file = $old['file'] }
    }
    $Index.entities[$EntityName] = $entry
}


function Write-DmfPullIndex {
    param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)]$Index)
    $doc = [ordered]@{
        schemaVersion  = $Index.schemaVersion
        environment    = $Index.environment
        environmentUrl = $Index.environmentUrl
        legalEntity    = $Index.legalEntity
        lastRun        = $Index.lastRun
        entities       = [ordered]@{}
    }
    foreach ($k in ($Index.entities.Keys | Sort-Object)) { $doc.entities[$k] = $Index.entities[$k] }
    Write-DmfJsonAtomic -Path (Join-Path $Folder '_pull.json') -Json ($doc | ConvertTo-Json -Depth 8)
}
