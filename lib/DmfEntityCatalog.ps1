<#
.SYNOPSIS
    Reads resources/entity-catalog.json, the shipped catalog of D365 F&O data
    entities.

.DESCRIPTION
    The catalog records, for every data entity in a given application version,
    the DMF label, the AOT name, whether the entity is reachable over OData,
    its collection name, key fields and company context.  It is committed to
    this repository as a data file so none of that has to be looked up at run
    time: the F&O Metadata service serves one label per call and refuses
    $filter on /Metadata/Labels with HTTP 501, which made resolving a large
    template slow, and it cannot describe a non-public entity at all.

    Dot-source this file to import:

        Get-DmfEntityCatalog          -- read the catalog JSON
        ConvertTo-DmfCatalogEntities  -- catalog document -> entity objects
        Get-DmfCatalogNormalizedLabel -- comparison form of a label
        New-DmfCatalogLabelLookup     -- label indexes for repeated lookups
        Find-DmfCatalogEntity         -- find an entity by its DMF label

    The catalog was produced from the AOT metadata of one application version
    (its sourceVersion field says which).  Regenerating it needs a development
    or build box; using it does not.  Treat the live Metadata service as the
    authority for any one environment: a stock catalog will not contain ISV or
    custom entities, and entity definitions change between versions.

.NOTES
    Write-Warn comes from DmfOutput.ps1 -- dot-source it first.
#>


function Get-DmfEntityCatalog {
    <#
    .SYNOPSIS
        Reads a catalog JSON; $null when the file is absent or unreadable.

    .OUTPUTS
        The parsed document: schemaVersion, sourceVersion, generatedAt,
        entityCount and entities (keyed by AOT entity name).
    #>
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    try { return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json) }
    catch {
        Write-Warn "Entity catalog at '$Path' could not be parsed: $($_.Exception.Message)"
        return $null
    }
}


function ConvertTo-DmfCatalogEntities {
    <#
    .SYNOPSIS
        Turns a catalog document into flat entity objects.

    .DESCRIPTION
        The file is keyed by AOT entity name to keep it compact; callers work
        with objects that carry the name alongside the rest.

    .OUTPUTS
        [pscustomobject[]] with Name, Label, IsPublic, DmEnabled, Category,
        PublicEntityName, Collection, CompanySpecific, RootTable, KeyFields.
        Wrap the call in @() before counting.
    #>
    # Not mandatory: a caller passing the result of Get-DmfEntityCatalog for a
    # missing file gets an empty list rather than a binding error.
    param($Document)

    if ($null -eq $Document) { return @() }
    $entitiesProperty = $Document.PSObject.Properties['entities']
    if ($null -eq $entitiesProperty -or $null -eq $entitiesProperty.Value) { return @() }

    $entities = [System.Collections.Generic.List[object]]::new()
    foreach ($property in $entitiesProperty.Value.PSObject.Properties) {
        $value = $property.Value
        $get = {
            param($name, $fallback)
            $found = $value.PSObject.Properties[$name]
            if ($null -ne $found -and $null -ne $found.Value) { $found.Value } else { $fallback }
        }
        $entities.Add([pscustomobject]@{
            Name             = $property.Name
            Label            = [string](& $get 'label' '')
            IsPublic         = [bool](& $get 'isPublic' $false)
            DmEnabled        = [bool](& $get 'dmEnabled' $false)
            Category         = [string](& $get 'category' '')
            PublicEntityName = [string](& $get 'publicEntityName' '')
            Collection       = [string](& $get 'collection' '')
            CompanySpecific  = [bool](& $get 'companySpecific' $false)
            RootTable        = [string](& $get 'rootTable' '')
            KeyFields        = @(& $get 'keyFields' @())
        })
    }
    return $entities.ToArray()
}


function Get-DmfCatalogNormalizedLabel {
    <#
    .SYNOPSIS
        Comparison form of a label: lower-cased, punctuation collapsed, version
        tokens dropped, crude singularisation, tokens sorted and de-duplicated.

    .DESCRIPTION
        Used to match a manifest's entity label against a catalog label when
        the two differ only cosmetically ('Country/regions' vs
        'Country/region').  Deliberately loose; callers try an exact label
        match first.
    #>
    param([AllowEmptyString()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $tokens = ($Text.ToLowerInvariant() -replace '[^a-z0-9]+', ' ').Trim() -split '\s+'
    $shaped = foreach ($token in $tokens) {
        if (-not $token) { continue }
        if ($token -match '^v\d+$') { continue }
        if ($token -match 'ies$')       { $token -replace 'ies$', 'y' }
        elseif ($token -match '[^s]s$') { $token -replace 's$', '' }
        else                            { $token }
    }
    return ((@($shaped) | Sort-Object -Unique) -join ' ')
}


function New-DmfCatalogLabelLookup {
    <#
    .SYNOPSIS
        Builds the exact and normalised label indexes Find-DmfCatalogEntity
        uses.  Build it once when resolving many labels.
    #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Entities)

    $exact      = @{}
    $normalized = @{}
    foreach ($entity in $Entities) {
        if (-not $entity.Label) { continue }
        if (-not $exact.ContainsKey($entity.Label)) { $exact[$entity.Label] = @() }
        $exact[$entity.Label] += $entity

        $form = Get-DmfCatalogNormalizedLabel -Text $entity.Label
        if (-not $form) { continue }
        if (-not $normalized.ContainsKey($form)) { $normalized[$form] = @() }
        $normalized[$form] += $entity
    }
    return [pscustomobject]@{ Exact = $exact; Normalized = $normalized }
}


function Find-DmfCatalogEntity {
    <#
    .SYNOPSIS
        Finds the catalog entity behind a DMF entity label.

    .DESCRIPTION
        Exact label match first, then the normalised form.  When several
        entities share a label the DMF-enabled one wins, because that is the
        one a template line refers to.

    .PARAMETER Entities
        Objects from ConvertTo-DmfCatalogEntities.  Ignored when -Index is given.

    .PARAMETER Index
        A prebuilt lookup from New-DmfCatalogLabelLookup.

    .OUTPUTS
        One entity object, or $null.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Label,
        [object[]]$Entities,
        $Index
    )

    if ([string]::IsNullOrWhiteSpace($Label)) { return $null }
    if ($null -eq $Index) {
        if ($null -eq $Entities) { return $null }
        $Index = New-DmfCatalogLabelLookup -Entities $Entities
    }

    $found = $null
    if ($Index.Exact.ContainsKey($Label)) { $found = $Index.Exact[$Label] }
    else {
        $normalized = Get-DmfCatalogNormalizedLabel -Text $Label
        if ($normalized -and $Index.Normalized.ContainsKey($normalized)) { $found = $Index.Normalized[$normalized] }
    }
    if ($null -eq $found) { return $null }

    $preferred = @($found | Where-Object { $_.DmEnabled })
    if ($preferred.Count -gt 0) { return $preferred[0] }
    return @($found)[0]
}
