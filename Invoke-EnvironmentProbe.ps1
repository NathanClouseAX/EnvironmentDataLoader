#Requires -Version 5.1
<#
.SYNOPSIS
    Probes a D365 F&O environment for the platform behaviours that the OData
    data pull (Invoke-ProjectExport.ps1 -Mode OData) and its entity resolution
    depend on, and writes the findings as JSON and Markdown.

.DESCRIPTION
    Runs six read-only checks and reports PASS / FAIL / INCONCLUSIVE for each:

      V1  /Metadata/DataEntities exposes Name, PublicEntityName,
          PublicCollectionName, LabelId, DataServiceEnabled,
          DataManagementEnabled and EntityCategory.  Also records the total
          row count and whether $filter is honoured.
      V2  /Metadata/PublicEntities exposes Properties[] with IsKey, and a
          dataAreaId property appears on the company-specific entity but not
          on the shared one.
      V3  /Metadata/Labels accepts an or-chained Id filter combined with a
          Language filter.
      V4  A DMF template line label (DefinitionGroupTemplateLines.Entity)
          matches the data entity's own label, so a label can be resolved to
          an AOT entity name.  Tries a reverse lookup by label text first
          (V4a) and a forward lookup from entity to template line second (V4b).
      V5  A company-specific collection needs cross-company=true together
          with $filter=dataAreaId eq '<LE>' to return that company's rows.
          Records what the same query returns without cross-company as a
          control.
      V6  Device-code sign-in with offline_access in the scope returns a
          refresh_token (needed for silent renewal on long runs).

    The script is read-only.  No access or refresh token is written to the
    output files or the transcript -- only their presence and length.

    Run it once per environment before relying on the OData path there; keep
    the Markdown with the environment's notes.  Any FAIL points at the
    behaviour the resolver will have to work around (see README, Entity
    resolution).

    Library files (in ./lib/)
    ─────────────────────────
    DmfOutput.ps1   -- Write-* helpers, Format-Elapsed, Stop-RunTranscript
    DmfRequest.ps1  -- Invoke-DmfRequest (retry + throttling), Get-DmfRetryAfterSeconds
    DmfAuth.ps1     -- Connect-DmfEnvironment (device code + offline_access), Get-DmfAuthHeaders

.PARAMETER EnvironmentUrl
    Base URL of the D365 F&O environment, e.g. https://contoso.operations.dynamics.com

.PARAMETER TenantId
    Microsoft Entra tenant ID or domain, e.g. contoso.onmicrosoft.com

.PARAMETER LegalEntityId
    Company used by the cross-company probe (V5).  Choose one that is NOT the
    signed-in user's default company and that has at least one row in
    -CompanyCollection, otherwise V5 cannot distinguish the two behaviours.

.PARAMETER SharedEntity
    Public entity name of an entity with no dataAreaId (V2).  Default: Currency.

.PARAMETER CompanyEntity
    Public entity name of a company-specific entity (V2).  Default: CustomerGroup.

.PARAMETER CompanyCollection
    OData collection (entity set) name for the company-specific entity (V5).
    Default: CustomerGroups.

.PARAMETER Language
    Label language for V3 / V4.  Default: en-US.

.PARAMETER SampleSize
    Number of rows to sample per probe (1-50).  Default: 5.

.PARAMETER OutputPath
    Directory for the JSON and Markdown result files.  Default: $env:TEMP.
    Pass '' to skip writing files (console only).

.PARAMETER AuthMode
    How to sign in.  Auto (default): open the default browser when the
    session is interactive, falling back to the device code flow if that
    fails; Browser: browser only; DeviceCode: print a code to enter in
    any browser (for SSH sessions and servers without a browser).

.PARAMETER LogPath
    Transcript path.  Auto-generated under -OutputPath (or $env:TEMP) when
    omitted; pass '' to suppress.

.PARAMETER MaxRetries
    Retry limit for transient REST failures.  Default 1 -- a probe should
    report quickly rather than mask a flaky endpoint behind retries.

.PARAMETER WhatIf
    List the requests that would be made and exit.  No authentication, no
    network calls.

.PARAMETER PassThru
    Emit one result object per probe to the pipeline:
    Id, Title, Result, Detail, Requests, Observed.

.EXAMPLE
    .\Invoke-EnvironmentProbe.ps1 `
        -EnvironmentUrl 'https://contoso-uat.sandbox.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'USMF'

.EXAMPLE
    # Keep the results with the environment's notes; use a different
    # company-specific entity because CustomerGroups is empty in this company
    .\Invoke-EnvironmentProbe.ps1 `
        -EnvironmentUrl    'https://contoso.operations.dynamics.com' `
        -TenantId          'contoso.onmicrosoft.com' `
        -LegalEntityId     'DEMF' `
        -CompanyEntity     'VendorGroup' `
        -CompanyCollection 'VendorGroups' `
        -OutputPath        'C:\DMF\probes'

.EXAMPLE
    # Preview the requests without signing in
    .\Invoke-EnvironmentProbe.ps1 -EnvironmentUrl 'https://contoso.operations.dynamics.com' -TenantId 'contoso.onmicrosoft.com' -LegalEntityId 'USMF' -WhatIf
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^https?://')]
    [string]$EnvironmentUrl,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$LegalEntityId,

    [ValidateNotNullOrEmpty()]
    [string]$SharedEntity = 'Currency',

    [ValidateNotNullOrEmpty()]
    [string]$CompanyEntity = 'CustomerGroup',

    [ValidateNotNullOrEmpty()]
    [string]$CompanyCollection = 'CustomerGroups',

    [ValidateNotNullOrEmpty()]
    [string]$Language = 'en-US',

    [ValidateRange(1, 50)]
    [int]$SampleSize = 5,

    [AllowEmptyString()]
    [string]$OutputPath = $env:TEMP,

    [ValidateSet('Auto', 'Browser', 'DeviceCode')]
    [string]$AuthMode = 'Auto',

    [AllowEmptyString()]
    [string]$LogPath,

    [ValidateRange(0, 10)]
    [int]$MaxRetries = 1,

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

# =============================================================================
#  Pre-flight validation
# =============================================================================
if ($OutputPath -ne '' -and -not (Test-Path $OutputPath -PathType Container)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# =============================================================================
#  Script-level constants  (consumed by lib functions via $Script: scope)
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
#  Helpers
# =============================================================================

function Get-EnvironmentNameFromUrl {
    <#
    .SYNOPSIS  Host's first DNS label, lower-cased, path-safe (same rule as Get-DmfEnvironmentName).
    #>
    param([Parameter(Mandatory)][string]$Url)
    $host1 = ([System.Uri]$Url).Host
    $label = ($host1 -split '\.')[0].ToLowerInvariant()
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() + @('.')
    foreach ($c in $invalid) { $label = $label.Replace([string]$c, '-') }
    return $label
}

function Test-HasProperty {
    param([Parameter(Mandatory)][AllowNull()]$Object, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $Object) { return $false }
    return ($null -ne $Object.PSObject.Properties[$Name])
}

function Get-PropertyOrNull {
    param([Parameter(Mandatory)][AllowNull()]$Object, [Parameter(Mandatory)][string]$Name)
    if (Test-HasProperty -Object $Object -Name $Name) { return $Object.$Name }
    return $null
}

function Get-ODataValue {
    <#
    .SYNOPSIS  Returns the 'value' array of an OData response (or the object itself when not a collection).
    #>
    param([Parameter(Mandatory)][AllowNull()]$Response)
    # The leading comma stops PowerShell unrolling the array on return, so a
    # zero- or one-row result is still an array and .Count / [0] work under
    # StrictMode on Windows PowerShell 5.1 (which has no intrinsic .Count on
    # scalars or $null in strict mode).
    if ($null -eq $Response) { return ,@() }
    # Only an array-valued 'value' is a collection; a single Labels entity has
    # a scalar 'Value' (its text) and must be returned whole.
    $prop = $Response.PSObject.Properties['value']
    if ($null -ne $prop -and ($prop.Value -is [array] -or ($prop.Value -is [System.Collections.IList] -and $prop.Value -isnot [string]))) { return ,@($prop.Value) }
    return ,@($Response)
}

function ConvertTo-ODataLiteral {
    param([Parameter(Mandatory)][string]$Value)
    return $Value.Replace("'", "''")
}

function New-ODataUri {
    param(
        [Parameter(Mandatory)][string]$Base,
        [Parameter(Mandatory)][string]$Path,
        [string]$Filter,
        [string]$Select,
        [int]$Top = 0,
        [switch]$Count,
        [switch]$CrossCompany
    )
    $parts = [System.Collections.Generic.List[string]]::new()
    if ($CrossCompany) { $parts.Add('cross-company=true') }
    if ($Filter)       { $parts.Add('$filter=' + [System.Uri]::EscapeDataString($Filter)) }
    if ($Select)       { $parts.Add('$select=' + [System.Uri]::EscapeDataString($Select)) }
    if ($Top -gt 0)    { $parts.Add("`$top=$Top") }
    if ($Count)        { $parts.Add('$count=true') }
    $uri = "$Base/$Path"
    if ($parts.Count -gt 0) { $uri += '?' + ($parts -join '&') }
    return $uri
}

# Every request made during a probe is recorded here (URL, status, timing)
# so the Markdown/JSON output shows exactly what was asked of the environment.
$Script:ProbeRequests = [System.Collections.Generic.List[pscustomobject]]::new()

function Invoke-ProbeGet {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$Operation
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $resp = Invoke-DmfRequest -Operation $Operation -Params @{
            Method  = 'Get'
            Uri     = $Uri
            Headers = $authHeaders
        }
        $sw.Stop()
        $Script:ProbeRequests.Add([pscustomobject]@{
            Operation = $Operation; Uri = $Uri; Status = 200
            ElapsedMs = $sw.ElapsedMilliseconds; Error = $null
        })
        return $resp
    }
    catch {
        $sw.Stop()
        $status = 0
        if ($_.Exception.Message -match 'HTTP (\d{3})') { $status = [int]$Matches[1] }
        $Script:ProbeRequests.Add([pscustomobject]@{
            Operation = $Operation; Uri = $Uri; Status = $status
            ElapsedMs = $sw.ElapsedMilliseconds; Error = $_.Exception.Message
        })
        throw
    }
}

function Invoke-Probe {
    <#
    .SYNOPSIS
        Runs one probe body, captures its requests and result, prints the outcome.
    .NOTES
        The body returns a hashtable: Result (PASS|FAIL|INCONCLUSIVE|SKIP),
        Detail (one line), Observed (small object for the report).
    #>
    param(
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Assumption,
        [Parameter(Mandatory)][scriptblock]$Body
    )

    Write-Step "$Id  $Title"
    $Script:ProbeRequests.Clear()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    $outcome = $null
    try {
        $outcome = & $Body
    }
    catch {
        $outcome = @{ Result = 'FAIL'; Detail = $_.Exception.Message; Observed = $null }
    }
    $sw.Stop()

    if ($null -eq $outcome -or -not ($outcome -is [hashtable])) {
        $outcome = @{ Result = 'INCONCLUSIVE'; Detail = 'Probe returned no outcome.'; Observed = $null }
    }
    foreach ($k in 'Result', 'Detail', 'Observed') {
        if (-not $outcome.ContainsKey($k)) { $outcome[$k] = $null }
    }

    $result = [pscustomobject]@{
        Id         = $Id
        Title      = $Title
        Assumption = $Assumption
        Result     = [string]$outcome.Result
        Detail     = [string]$outcome.Detail
        ElapsedMs  = $sw.ElapsedMilliseconds
        Requests   = @($Script:ProbeRequests.ToArray())
        Observed   = $outcome.Observed
    }

    switch ($result.Result) {
        'PASS'         { Write-OK   "$Id PASS  -- $($result.Detail)" }
        'FAIL'         { Write-Fail "$Id FAIL  -- $($result.Detail)" }
        'SKIP'         { Write-Info "$Id SKIP  -- $($result.Detail)" }
        default        { Write-Warn "$Id INCONCLUSIVE  -- $($result.Detail)" }
    }
    return $result
}

function Limit-Sample {
    <#
    .SYNOPSIS  Keeps only the named properties of up to N objects, for the report.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Items,
        [Parameter(Mandatory)][string[]]$Properties,
        [int]$Max = 5
    )
    $out = [System.Collections.Generic.List[pscustomobject]]::new()
    foreach ($item in ($Items | Select-Object -First $Max)) {
        $o = [ordered]@{}
        foreach ($p in $Properties) { $o[$p] = Get-PropertyOrNull -Object $item -Name $p }
        $out.Add([pscustomobject]$o)
    }
    return ,$out.ToArray()
}

# =============================================================================
#  1.  Transcript startup
# =============================================================================
$scriptStart = Get-Date
$ts          = Get-Date -Format 'yyyyMMdd_HHmmss'
$baseUrl     = $EnvironmentUrl.TrimEnd('/')
$envName     = Get-EnvironmentNameFromUrl -Url $baseUrl
$logBase     = if ($OutputPath -ne '') { $OutputPath } else { $env:TEMP }

if (-not $PSBoundParameters.ContainsKey('LogPath')) {
    $LogPath = Join-Path $logBase "EnvironmentProbes_${envName}_${ts}.log"
}
if ($LogPath -ne '' -and -not $WhatIf) {
    try {
        Start-Transcript -Path $LogPath -Force | Out-Null
        $Script:TranscriptActive = $true
    } catch { <# transcript not supported in this host -- continue silently #> }
}

# =============================================================================
#  2.  Banner + run summary
# =============================================================================
Write-Banner -DryRun:$WhatIf -Title 'D365 F&O Environment Probe'
Write-Info "Environment  : $EnvironmentUrl  ($envName)"
Write-Info "Tenant       : $TenantId"
Write-Info "Legal entity : $LegalEntityId  (cross-company probe)"
Write-Info "Entities     : shared=$SharedEntity  company=$CompanyEntity  collection=$CompanyCollection"
Write-Info "Language     : $Language"
Write-Info "Sample size  : $SampleSize"
Write-Info "Output       : $(if ($OutputPath -ne '') { $OutputPath } else { '(console only)' })"
if ($LogPath -ne '') { Write-Info "Log          : $LogPath" }

# =============================================================================
#  3.  Request plan  (also the -WhatIf output)
# =============================================================================
$metaBase = "$baseUrl/Metadata"
$dataBase = "$baseUrl/data"
$ci       = [System.StringComparison]::OrdinalIgnoreCase

$plan = @(
    [pscustomobject]@{ Id = 'V1'; Uri = New-ODataUri -Base $metaBase -Path 'DataEntities' -Top $SampleSize }
    [pscustomobject]@{ Id = 'V1'; Uri = New-ODataUri -Base $metaBase -Path 'DataEntities' -Top 1 -Count }
    [pscustomobject]@{ Id = 'V1'; Uri = New-ODataUri -Base $metaBase -Path 'DataEntities' -Filter 'DataManagementEnabled eq true' -Top $SampleSize }
    [pscustomobject]@{ Id = 'V2'; Uri = New-ODataUri -Base $metaBase -Path 'PublicEntities' -Filter "Name eq '$(ConvertTo-ODataLiteral $SharedEntity)'" }
    [pscustomobject]@{ Id = 'V2'; Uri = New-ODataUri -Base $metaBase -Path 'PublicEntities' -Filter "Name eq '$(ConvertTo-ODataLiteral $CompanyEntity)'" }
    [pscustomobject]@{ Id = 'V3'; Uri = New-ODataUri -Base $metaBase -Path 'Labels' -Filter "(Id eq '<id1>' or Id eq '<id2>') and Language eq '$Language'" }
    [pscustomobject]@{ Id = 'V4'; Uri = New-ODataUri -Base $dataBase -Path 'DefinitionGroupTemplateLines' -Top $SampleSize }
    [pscustomobject]@{ Id = 'V4'; Uri = New-ODataUri -Base $metaBase -Path 'Labels' -Filter "Value eq '<template line label>' and Language eq '$Language'" }
    [pscustomobject]@{ Id = 'V4'; Uri = New-ODataUri -Base $metaBase -Path 'DataEntities' -Filter "LabelId eq '<label id>'" }
    [pscustomobject]@{ Id = 'V5'; Uri = New-ODataUri -Base $dataBase -Path $CompanyCollection -CrossCompany -Filter "dataAreaId eq '$(ConvertTo-ODataLiteral $LegalEntityId)'" -Top $SampleSize }
    [pscustomobject]@{ Id = 'V5'; Uri = New-ODataUri -Base $dataBase -Path $CompanyCollection -Filter "dataAreaId eq '$(ConvertTo-ODataLiteral $LegalEntityId)'" -Top $SampleSize }
    [pscustomobject]@{ Id = 'V5'; Uri = New-ODataUri -Base $dataBase -Path $CompanyCollection -Top 1 }
    [pscustomobject]@{ Id = 'V6'; Uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode  (scope: $baseUrl/.default offline_access)" }
)

if ($WhatIf) {
    Write-Rule 'WhatIf -- requests that would be made'
    foreach ($p in $plan) { Write-Info ("{0}  GET {1}" -f $p.Id, $p.Uri) }
    Write-Host ''
    Write-Info 'No API calls made.  Remove -WhatIf to run the probes.'
    Stop-RunTranscript
    return
}

# =============================================================================
#  4.  Authenticate  (device code flow; offline_access requested for V6)
# =============================================================================
Write-Step 'Authenticating with Microsoft Entra (device code flow)'

$session = Connect-DmfEnvironment -EnvironmentUrl $baseUrl -TenantId $TenantId -AuthMode $AuthMode
$Script:DmfSession = $session
$authHeaders       = Get-DmfAuthHeaders -Session $session

# V6 evidence -- presence and length only; the token itself is never recorded.
$v6Observed = [ordered]@{
    ScopeRequested      = "$baseUrl/.default offline_access"
    ScopeGranted        = $session.Scope
    RefreshTokenPresent = (-not [string]::IsNullOrEmpty($session.RefreshToken))
    RefreshTokenLength  = $(if ($session.RefreshToken) { $session.RefreshToken.Length } else { 0 })
    ExpiresAt           = $session.ExpiresAt.ToString('o')
    Note                = $session.ScopeNote
}

# =============================================================================
#  5.  Probes
# =============================================================================
$results       = [System.Collections.Generic.List[pscustomobject]]::new()
$v1Sample      = @()   # shared with V3 / V4b
$v1DmSample    = @()

# ---------------------------------------------------------------------------
$results.Add((Invoke-Probe -Id 'V1' -Title 'Metadata service: DataEntities shape' `
    -Assumption 'GET /Metadata/DataEntities returns Name, PublicEntityName, PublicCollectionName, LabelId, DataServiceEnabled, DataManagementEnabled, EntityCategory' `
    -Body {
        $expected = 'Name', 'PublicEntityName', 'PublicCollectionName', 'LabelId', 'DataServiceEnabled', 'DataManagementEnabled', 'EntityCategory'

        $resp  = Invoke-ProbeGet -Operation 'DataEntities sample' -Uri (New-ODataUri -Base $metaBase -Path 'DataEntities' -Top $SampleSize)
        $rows  = Get-ODataValue -Response $resp
        if ($rows.Count -eq 0) { return @{ Result = 'FAIL'; Detail = 'Endpoint answered but returned no rows.'; Observed = $null } }
        $Script:v1Sample = $rows

        $present = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
        $missing = @($expected | Where-Object { $present -notcontains $_ })
        $extra   = @($present  | Where-Object { $expected -notcontains $_ -and $_ -notlike '@odata*' })

        # Total row count (informational -- the resolver fetches this list once per run)
        $total = $null
        try {
            $cResp = Invoke-ProbeGet -Operation 'DataEntities count' -Uri (New-ODataUri -Base $metaBase -Path 'DataEntities' -Top 1 -Count)
            $total = Get-PropertyOrNull -Object $cResp -Name '@odata.count'
        } catch { Write-Detail "  `$count not supported: $($_.Exception.Message)" }

        # $filter support (needed for V4 and for cheap single-entity lookups)
        $filterOk = $false
        $dmRows   = @()
        try {
            $fResp = Invoke-ProbeGet -Operation 'DataEntities filter' -Uri (New-ODataUri -Base $metaBase -Path 'DataEntities' -Filter 'DataManagementEnabled eq true' -Top $SampleSize)
            $dmRows = Get-ODataValue -Response $fResp
            $filterOk = ($dmRows.Count -gt 0) -and (@($dmRows | Where-Object { -not (Get-PropertyOrNull $_ 'DataManagementEnabled') }).Count -eq 0)
            $Script:v1DmSample = $dmRows
        } catch { Write-Detail "  `$filter failed: $($_.Exception.Message)" }

        $observed = [ordered]@{
            PropertiesPresent = $present
            MissingExpected   = $missing
            ExtraProperties   = $extra
            TotalRows         = $total
            FilterSupported   = $filterOk
            Sample            = Limit-Sample -Items $rows -Properties $expected -Max $SampleSize
        }

        if ($missing.Count -gt 0) {
            return @{ Result = 'FAIL'; Detail = "Missing expected properties: $($missing -join ', ')"; Observed = $observed }
        }
        $countTxt  = if ($null -ne $total) { "$total rows total" } else { 'row count unknown ($count unsupported)' }
        $filterTxt = if ($filterOk) { '$filter honoured' } else { '$filter NOT honoured' }
        return @{ Result = 'PASS'; Detail = "All 7 expected properties present; $countTxt; $filterTxt"; Observed = $observed }
    }))

# ---------------------------------------------------------------------------
$results.Add((Invoke-Probe -Id 'V2' -Title 'Metadata service: PublicEntities keys and dataAreaId' `
    -Assumption 'GET /Metadata/PublicEntities returns Properties[] with IsKey; dataAreaId is present only on company-specific entities' `
    -Body {
        $inspect = {
            param($name)
            $resp = Invoke-ProbeGet -Operation "PublicEntities $name" -Uri (New-ODataUri -Base $metaBase -Path 'PublicEntities' -Filter "Name eq '$(ConvertTo-ODataLiteral $name)'")
            $rows = Get-ODataValue -Response $resp
            if ($rows.Count -eq 0) { return [ordered]@{ Name = $name; Found = $false } }
            $e     = $rows[0]
            $props = @(Get-PropertyOrNull -Object $e -Name 'Properties')
            $keys  = @($props | Where-Object { Get-PropertyOrNull $_ 'IsKey' } | ForEach-Object { Get-PropertyOrNull $_ 'Name' })
            $names = @($props | ForEach-Object { Get-PropertyOrNull $_ 'Name' })
            return [ordered]@{
                Name            = $name
                Found           = $true
                EntitySetName   = Get-PropertyOrNull -Object $e -Name 'EntitySetName'
                LabelId         = Get-PropertyOrNull -Object $e -Name 'LabelId'
                PropertyCount   = $props.Count
                HasIsKeyFlag    = ($props.Count -gt 0 -and (Test-HasProperty -Object $props[0] -Name 'IsKey'))
                KeyFields       = $keys
                HasDataAreaId   = ($names -contains 'dataAreaId')
                TopLevelFields  = @($e.PSObject.Properties | ForEach-Object { $_.Name } | Where-Object { $_ -notlike '@odata*' })
            }
        }
        $shared  = & $inspect $SharedEntity
        $company = & $inspect $CompanyEntity
        $observed = [ordered]@{ Shared = $shared; Company = $company }

        if (-not $shared.Found -or -not $company.Found) {
            $which = @(@($shared, $company) | Where-Object { -not $_.Found } | ForEach-Object { $_.Name }) -join ', '
            return @{ Result = 'INCONCLUSIVE'; Detail = "Public entity not found: $which.  Pass -SharedEntity / -CompanyEntity with names that exist here."; Observed = $observed }
        }
        if (-not $shared.HasIsKeyFlag -or -not $company.HasIsKeyFlag) {
            return @{ Result = 'FAIL'; Detail = 'Properties[] has no IsKey flag.'; Observed = $observed }
        }
        if ($shared.KeyFields.Count -eq 0 -or $company.KeyFields.Count -eq 0) {
            return @{ Result = 'FAIL'; Detail = 'IsKey present but no key fields flagged on at least one entity.'; Observed = $observed }
        }
        if ($shared.HasDataAreaId) {
            return @{ Result = 'INCONCLUSIVE'; Detail = "$SharedEntity exposes dataAreaId; pick a truly shared entity via -SharedEntity to confirm the negative case."; Observed = $observed }
        }
        if (-not $company.HasDataAreaId) {
            return @{ Result = 'FAIL'; Detail = "$CompanyEntity does not expose dataAreaId -- company detection by property would not work."; Observed = $observed }
        }
        return @{ Result = 'PASS'; Detail = "Keys: $SharedEntity=[$($shared.KeyFields -join ',')]  $CompanyEntity=[$($company.KeyFields -join ',')]; dataAreaId only on $CompanyEntity"; Observed = $observed }
    }))

# ---------------------------------------------------------------------------
$results.Add((Invoke-Probe -Id 'V3' -Title 'Metadata service: Labels or-chained Id filter' `
    -Assumption "GET /Metadata/Labels accepts (Id eq 'a' or Id eq 'b') and Language eq '$Language'" `
    -Body {
        $ids = @($Script:v1DmSample + $Script:v1Sample |
                 ForEach-Object { Get-PropertyOrNull $_ 'LabelId' } |
                 Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                 Select-Object -Unique -First 3)
        if ($ids.Count -lt 2) {
            return @{ Result = 'SKIP'; Detail = 'Fewer than two LabelIds available from V1 to test with.'; Observed = $null }
        }

        $idFilter = ($ids | ForEach-Object { "Id eq '$(ConvertTo-ODataLiteral $_)'" }) -join ' or '
        $uri      = New-ODataUri -Base $metaBase -Path 'Labels' -Filter "($idFilter) and Language eq '$(ConvertTo-ODataLiteral $Language)'"

        # Form 1: or-chained $filter (cheap batches, if the service allows it)
        $batchRows = @()
        $batchErr  = $null
        try {
            $batchRows = Get-ODataValue -Response (Invoke-ProbeGet -Operation 'Labels batch' -Uri $uri)
        } catch { $batchErr = $_.Exception.Message }

        # Form 2: key segment Labels(Id='...',Language='...') -- what the resolver relies on
        $keyRows = @()
        $keyErr  = $null
        try {
            $kUri = "$metaBase/Labels(Id='$([System.Uri]::EscapeDataString((ConvertTo-ODataLiteral $ids[0])))',Language='$([System.Uri]::EscapeDataString((ConvertTo-ODataLiteral $Language)))')"
            $keyRows = Get-ODataValue -Response (Invoke-ProbeGet -Operation 'Labels key segment' -Uri $kUri)
        } catch { $keyErr = $_.Exception.Message }

        $observed = [ordered]@{
            IdsTested        = $ids
            FilterBatchCount = $batchRows.Count
            FilterBatchError = $batchErr
            FilterValues     = Limit-Sample -Items $batchRows -Properties 'Id', 'Language', 'Value' -Max 10
            KeySegmentValue  = Limit-Sample -Items $keyRows -Properties 'Id', 'Language', 'Value' -Max 1
            KeySegmentError  = $keyErr
        }

        if ($keyErr) {
            return @{ Result = 'FAIL'; Detail = "Key-segment label read failed ($keyErr)$(if (-not $batchErr) { '; the or-chained filter works' })"; Observed = $observed }
        }
        if ($batchErr) {
            return @{ Result = 'PASS'; Detail = "Key-segment reads work (one label per call); the or-chained `$filter is refused ($batchErr)"; Observed = $observed }
        }
        return @{ Result = 'PASS'; Detail = "Both forms work: or-chain returned $($batchRows.Count) label(s) for $($ids.Count) id(s), key segment returned $($keyRows.Count)"; Observed = $observed }
    }))

# ---------------------------------------------------------------------------
$results.Add((Invoke-Probe -Id 'V4' -Title 'DMF template label equals data-entity label' `
    -Assumption 'DefinitionGroupTemplateLines.Entity equals the resolved label of a DataEntities row, so labels can be mapped to AOT names' `
    -Body {
        $linesResp = Invoke-ProbeGet -Operation 'template lines sample' -Uri (New-ODataUri -Base $dataBase -Path 'DefinitionGroupTemplateLines' -Top $SampleSize)
        $lines     = Get-ODataValue -Response $linesResp
        if ($lines.Count -eq 0) {
            return @{ Result = 'INCONCLUSIVE'; Detail = 'No template lines in this environment.  Load default templates (Data management > Templates) and re-run.'; Observed = $null }
        }
        $labels = @($lines | ForEach-Object { Get-PropertyOrNull $_ 'Entity' } | Where-Object { $_ } | Select-Object -Unique)

        # -- V4a: reverse lookup by label text -----------------------------------
        $perLabel     = [System.Collections.Generic.List[pscustomobject]]::new()
        $valueFilterOk = $true
        $valueFilterErr = $null
        foreach ($label in $labels) {
            $row = [ordered]@{ Label = $label; LabelIds = @(); Entities = @(); Resolved = $false }
            try {
                $lUri  = New-ODataUri -Base $metaBase -Path 'Labels' -Filter "Value eq '$(ConvertTo-ODataLiteral $label)' and Language eq '$(ConvertTo-ODataLiteral $Language)'"
                $lRows = Get-ODataValue -Response (Invoke-ProbeGet -Operation "Labels by value '$label'" -Uri $lUri)
                $row.LabelIds = @($lRows | ForEach-Object { Get-PropertyOrNull $_ 'Id' } | Where-Object { $_ } | Select-Object -Unique -First 10)
            } catch {
                $valueFilterOk  = $false
                $valueFilterErr = $_.Exception.Message
                $perLabel.Add([pscustomobject]$row)
                break
            }
            if ($row.LabelIds.Count -gt 0) {
                $idFilter = ($row.LabelIds | ForEach-Object { "LabelId eq '$(ConvertTo-ODataLiteral $_)'" }) -join ' or '
                try {
                    $eRows = Get-ODataValue -Response (Invoke-ProbeGet -Operation "DataEntities by LabelId '$label'" -Uri (New-ODataUri -Base $metaBase -Path 'DataEntities' -Filter $idFilter))
                    $row.Entities = @($eRows | ForEach-Object {
                        [pscustomobject]@{
                            Name                  = Get-PropertyOrNull $_ 'Name'
                            PublicCollectionName  = Get-PropertyOrNull $_ 'PublicCollectionName'
                            DataManagementEnabled = Get-PropertyOrNull $_ 'DataManagementEnabled'
                            DataServiceEnabled    = Get-PropertyOrNull $_ 'DataServiceEnabled'
                        }
                    })
                    $row.Resolved = ($row.Entities.Count -gt 0)
                } catch { Write-Detail "  DataEntities LabelId filter failed: $($_.Exception.Message)" }
            }
            $perLabel.Add([pscustomobject]$row)
        }

        if ($valueFilterOk) {
            $resolved  = @($perLabel | Where-Object { $_.Resolved }).Count
            $ambiguous = @($perLabel | Where-Object { @($_.Entities | Where-Object { $_.DataManagementEnabled }).Count -gt 1 }).Count
            $observed  = [ordered]@{ Method = 'V4a reverse lookup (Labels Value filter)'; Labels = $perLabel.ToArray(); ValueFilterSupported = $true }
            if ($resolved -eq $labels.Count) {
                return @{ Result = 'PASS'; Detail = "$resolved/$($labels.Count) template labels resolved to a DataEntities row ($ambiguous ambiguous)"; Observed = $observed }
            }
            if ($resolved -gt 0) {
                return @{ Result = 'INCONCLUSIVE'; Detail = "$resolved/$($labels.Count) resolved; unresolved labels need the entity-map fallback"; Observed = $observed }
            }
            return @{ Result = 'FAIL'; Detail = 'Labels Value filter works but no template label matched a DataEntities label.'; Observed = $observed }
        }

        # -- V4b: forward check -- entity label -> template line ------------------
        Write-Detail "  Labels Value filter not usable ($valueFilterErr); trying forward lookup."
        $forward = [System.Collections.Generic.List[pscustomobject]]::new()
        $candidates = @($Script:v1DmSample | Where-Object { Get-PropertyOrNull $_ 'LabelId' } | Select-Object -First $SampleSize)
        foreach ($ent in $candidates) {
            $labelId = Get-PropertyOrNull $ent 'LabelId'
            $value   = $null
            try {
                $kUri  = "$metaBase/Labels(Id='$([System.Uri]::EscapeDataString((ConvertTo-ODataLiteral $labelId)))',Language='$([System.Uri]::EscapeDataString((ConvertTo-ODataLiteral $Language)))')"
                $lRows = Get-ODataValue -Response (Invoke-ProbeGet -Operation "Label $labelId" -Uri $kUri)
                if ($lRows.Count -gt 0) { $value = Get-PropertyOrNull $lRows[0] 'Value' }
            } catch {}
            $inTemplate = $false
            if ($value) {
                try {
                    $tRows = Get-ODataValue -Response (Invoke-ProbeGet -Operation "template line '$value'" -Uri (New-ODataUri -Base $dataBase -Path 'DefinitionGroupTemplateLines' -Filter "Entity eq '$(ConvertTo-ODataLiteral $value)'" -Top 1))
                    $inTemplate = ($tRows.Count -gt 0)
                } catch {}
            }
            $forward.Add([pscustomobject]@{ Entity = (Get-PropertyOrNull $ent 'Name'); LabelId = $labelId; LabelValue = $value; FoundInTemplateLines = $inTemplate })
        }
        $hits = @($forward | Where-Object { $_.FoundInTemplateLines }).Count
        $observed = [ordered]@{ Method = 'V4b forward lookup'; ValueFilterSupported = $false; ValueFilterError = $valueFilterErr; Entities = $forward.ToArray() }
        if ($hits -gt 0) {
            return @{ Result = 'PASS'; Detail = "Value filter unsupported, but $hits/$($forward.Count) sampled entity labels (read by key segment) appear verbatim as template lines -- labels do match, so name-candidate resolution works"; Observed = $observed }
        }
        return @{ Result = 'INCONCLUSIVE'; Detail = "Value filter unsupported and none of $($forward.Count) sampled entities appear in template lines (sample may simply not be in any template)"; Observed = $observed }
    }))

# ---------------------------------------------------------------------------
$results.Add((Invoke-Probe -Id 'V5' -Title 'cross-company + dataAreaId filter' `
    -Assumption "GET /data/${CompanyCollection}?cross-company=true&`$filter=dataAreaId eq '$LegalEntityId' returns rows for $LegalEntityId" `
    -Body {
        $le = $LegalEntityId

        $ccRows = Get-ODataValue -Response (Invoke-ProbeGet -Operation 'cross-company + filter' -Uri (New-ODataUri -Base $dataBase -Path $CompanyCollection -CrossCompany -Filter "dataAreaId eq '$(ConvertTo-ODataLiteral $le)'" -Top $SampleSize))
        $ccAreas = @($ccRows | ForEach-Object { Get-PropertyOrNull $_ 'dataAreaId' })

        # Control 1: same filter without cross-company
        $noCcRows = @(); $noCcErr = $null
        try {
            $noCcRows = Get-ODataValue -Response (Invoke-ProbeGet -Operation 'filter only (no cross-company)' -Uri (New-ODataUri -Base $dataBase -Path $CompanyCollection -Filter "dataAreaId eq '$(ConvertTo-ODataLiteral $le)'" -Top $SampleSize))
        } catch { $noCcErr = $_.Exception.Message }

        # Control 2: no filter, no cross-company -> reveals the caller's default company
        $defRows = @(); $defErr = $null
        try {
            $defRows = Get-ODataValue -Response (Invoke-ProbeGet -Operation 'default company (no params)' -Uri (New-ODataUri -Base $dataBase -Path $CompanyCollection -Top 1))
        } catch { $defErr = $_.Exception.Message }
        $defaultCompany = if ($defRows.Count -gt 0) { Get-PropertyOrNull $defRows[0] 'dataAreaId' } else { $null }

        $observed = [ordered]@{
            Collection                  = $CompanyCollection
            RequestedCompany            = $le
            CrossCompanyRows            = $ccRows.Count
            CrossCompanyDataAreaIds     = @($ccAreas | Select-Object -Unique)
            FilterOnlyRows              = $noCcRows.Count
            FilterOnlyDataAreaIds       = @($noCcRows | ForEach-Object { Get-PropertyOrNull $_ 'dataAreaId' } | Select-Object -Unique)
            FilterOnlyError             = $noCcErr
            DefaultCompanyObserved      = $defaultCompany
            DefaultCompanyError         = $defErr
        }

        if ($ccRows.Count -eq 0) {
            return @{ Result = 'INCONCLUSIVE'; Detail = "$CompanyCollection has no rows in $le (or the entity is not company-specific).  Re-run with -LegalEntityId / -CompanyCollection that have data."; Observed = $observed }
        }
        $wrong = @($ccAreas | Where-Object { -not [string]::Equals([string]$_, $le, $ci) })
        if ($wrong.Count -gt 0) {
            return @{ Result = 'FAIL'; Detail = "cross-company query returned rows for other companies: $($wrong | Select-Object -Unique -join ', ')"; Observed = $observed }
        }
        $ctrl = if ($null -ne $defaultCompany -and -not [string]::Equals([string]$defaultCompany, $le, $ci)) {
            "default company is $defaultCompany; filter-only returned $($noCcRows.Count) row(s)"
        } elseif ($null -ne $defaultCompany) {
            "default company is also $le, so the control cannot show the difference -- re-run with a different -LegalEntityId for a stronger result"
        } else { 'default-company control unavailable' }
        return @{ Result = 'PASS'; Detail = "cross-company returned $($ccRows.Count) row(s), all dataAreaId=$le; $ctrl"; Observed = $observed }
    }))

# ---------------------------------------------------------------------------
$results.Add((Invoke-Probe -Id 'V6' -Title 'Device-code sign-in returns a refresh_token' `
    -Assumption "Requesting '$baseUrl/.default offline_access' yields refresh_token for client $($session.ClientId)" `
    -Body {
        $obs = $v6Observed
        if ($obs.Note) {
            return @{ Result = 'FAIL'; Detail = $obs.Note; Observed = $obs }
        }
        if ($obs.RefreshTokenPresent) {
            return @{ Result = 'PASS'; Detail = "refresh_token present (length $($obs.RefreshTokenLength)); access token valid until $($obs.ExpiresAt)"; Observed = $obs }
        }
        return @{ Result = 'FAIL'; Detail = 'Sign-in succeeded but no refresh_token in the token response.'; Observed = $obs }
    }))

# =============================================================================
#  6.  Summary
# =============================================================================
Write-Rule 'Summary'
Write-Host ("  {0,-4} {1,-13} {2,-46} {3}" -f 'Id', 'Result', 'Probe', 'Elapsed') -ForegroundColor White
foreach ($r in $results) {
    $colour = switch ($r.Result) { 'PASS' { 'Green' } 'FAIL' { 'Red' } 'SKIP' { 'DarkGray' } default { 'Yellow' } }
    $title  = if ($r.Title.Length -gt 46) { $r.Title.Substring(0, 43) + '...' } else { $r.Title }
    Write-Host ("  {0,-4} {1,-13} {2,-46} {3}" -f $r.Id, $r.Result, $title, (Format-Elapsed ([TimeSpan]::FromMilliseconds($r.ElapsedMs)))) -ForegroundColor $colour
}
Write-Host ''
$passCount = @($results | Where-Object Result -eq 'PASS').Count
$failCount = @($results | Where-Object Result -eq 'FAIL').Count
$otherCount = $results.Count - $passCount - $failCount
Write-Info "PASS $passCount   FAIL $failCount   INCONCLUSIVE/SKIP $otherCount   in $(Format-Elapsed ((Get-Date) - $scriptStart))"

# =============================================================================
#  7.  Write JSON + Markdown
# =============================================================================
if ($OutputPath -ne '') {
    $report = [ordered]@{
        environmentUrl  = $baseUrl
        environmentName = $envName
        tenantId        = $TenantId
        legalEntityId   = $LegalEntityId
        ranAt           = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        psVersion       = $PSVersionTable.PSVersion.ToString()
        scriptVersion   = $Script:Version
        parameters      = [ordered]@{ SharedEntity = $SharedEntity; CompanyEntity = $CompanyEntity; CompanyCollection = $CompanyCollection; Language = $Language; SampleSize = $SampleSize }
        probes          = @($results | ForEach-Object {
            [ordered]@{
                id = $_.Id; title = $_.Title; assumption = $_.Assumption; result = $_.Result; detail = $_.Detail
                elapsedMs = $_.ElapsedMs; requests = $_.Requests; observed = $_.Observed
            }
        })
    }

    $jsonPath = Join-Path $OutputPath "EnvironmentProbes_${envName}_${ts}.json"
    $report | ConvertTo-Json -Depth 12 | Set-Content -Path $jsonPath -Encoding UTF8
    Write-Info "JSON     : $jsonPath"

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("# Environment probe results: $envName")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine('| | |')
    [void]$sb.AppendLine('|---|---|')
    [void]$sb.AppendLine("| Environment | ``$baseUrl`` |")
    [void]$sb.AppendLine("| Legal entity (V5) | ``$LegalEntityId`` |")
    [void]$sb.AppendLine("| Ran at | $($report.ranAt) |")
    [void]$sb.AppendLine("| PowerShell | $($report.psVersion) |")
    [void]$sb.AppendLine("| Script | Invoke-EnvironmentProbe.ps1 v$($Script:Version) |")
    [void]$sb.AppendLine("| Parameters | shared=``$SharedEntity`` company=``$CompanyEntity`` collection=``$CompanyCollection`` language=``$Language`` sample=$SampleSize |")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine('## Summary')
    [void]$sb.AppendLine()
    [void]$sb.AppendLine('| Probe | Result | Detail |')
    [void]$sb.AppendLine('|---|---|---|')
    foreach ($r in $results) {
        [void]$sb.AppendLine("| $($r.Id) $($r.Title) | **$($r.Result)** | $($r.Detail -replace '\|', '\|') |")
    }
    foreach ($r in $results) {
        [void]$sb.AppendLine()
        [void]$sb.AppendLine("## $($r.Id) - $($r.Title)")
        [void]$sb.AppendLine()
        [void]$sb.AppendLine("**Assumption:** $($r.Assumption)")
        [void]$sb.AppendLine()
        [void]$sb.AppendLine("**Result:** $($r.Result) - $($r.Detail)")
        [void]$sb.AppendLine()
        if ($r.Requests.Count -gt 0) {
            [void]$sb.AppendLine('| Request | HTTP | ms |')
            [void]$sb.AppendLine('|---|---|---|')
            foreach ($q in $r.Requests) {
                $status = if ($q.Status -gt 0) { $q.Status } else { 'n/a' }
                $err    = if ($q.Error) { " - $($q.Error -replace '\|', '\|')" } else { '' }
                [void]$sb.AppendLine("| ``$($q.Uri)`` | $status$err | $($q.ElapsedMs) |")
            }
            [void]$sb.AppendLine()
        }
        if ($null -ne $r.Observed) {
            [void]$sb.AppendLine('<details><summary>Observed</summary>')
            [void]$sb.AppendLine()
            [void]$sb.AppendLine('```json')
            [void]$sb.AppendLine(($r.Observed | ConvertTo-Json -Depth 10))
            [void]$sb.AppendLine('```')
            [void]$sb.AppendLine()
            [void]$sb.AppendLine('</details>')
        }
    }

    $mdPath = Join-Path $OutputPath "EnvironmentProbes_${envName}_${ts}.md"
    Set-Content -Path $mdPath -Value $sb.ToString() -Encoding UTF8
    Write-Info "Markdown : $mdPath"
    Write-Info 'Keep the Markdown with this environment''s notes; any FAIL above is a behaviour the OData path must work around.'
}

Stop-RunTranscript

if ($PassThru) { $results | Write-Output }
