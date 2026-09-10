#Requires -Version 5.1
<#
.SYNOPSIS
    Builds DMF export projects from selected templates and exports them.

.DESCRIPTION
    1. Authenticates once via Microsoft Entra device code flow.
    2. Lists templates: local folders under -ResourcesPath that hold a
       Manifest.xml (-TemplateSource Local, the default) or the environment's
       validated DMF templates (-TemplateSource Environment).
    3. Presents an interactive selection menu (skipped when -TemplateName or -All
       is supplied).
    4. Prompts for the source legal entity (skipped when -LegalEntityId is supplied).
    5. -Mode Dmf (the default) -- for each selected template:
       a. Reads the entity lines (local Manifest.xml, or DefinitionGroupTemplateLines).
       b. Deletes any existing DMF project named "<TemplateId> <LegalEntityId>".
       c. Creates a fresh export project (DataManagementDefinitionGroups).
       d. Adds one entity record per template line (DataManagementDefinitionGroupDetails).
       e. Submits an ExportToPackage job to D365 DMF.
       f. Polls until the job reaches a terminal status.
       g. Retrieves the blob download URL and downloads the package to -DownloadPath.
    6. Prints a colour-coded summary table with per-template elapsed time and
       execution IDs for follow-up in D365 Job history.
    7. Writes a timestamped transcript to -LogPath (auto-generated when omitted;
       pass an empty string to suppress transcript logging entirely).

    -Mode OData reads the same template entities straight from the OData
    endpoints instead of running a DMF export job.  Nothing is created in
    D365.  The entities of every selected template are combined (each entity
    once), resolved to their OData collections through resources/entity-map.json
    and the F&O Metadata service, and written as JSON under
    -DataPath/<environment>/<legal entity>/<Entity label>.json together with a
    _pull.json run index.  Company-specific entities are read with
    cross-company=true and a dataAreaId filter.  Entities that are not
    OData-enabled or cannot be resolved are reported and skipped.  The output
    folder is the input to Compare-EnvironmentData.ps1.

    -WhatIf note:
      When -TemplateName is also supplied, no API calls are made at all.
      Otherwise the template list is fetched (read-only) to populate the menu
      (or to resolve -All); no projects are created or exported.

    All REST calls include automatic retry with exponential back-off and jitter
    (HTTP 5xx / 408 / transient network errors).  HTTP 429 throttling is handled
    separately: the server's Retry-After hint is honoured exactly and draws on
    its own retry budget, so being throttled does not consume the allowance
    reserved for genuine transient faults.  The token expiry window is tracked
    and surfaced as a warning when approaching expiry.

    Library files (in ./lib/)
    ─────────────────────────
    DmfOutput.ps1   -- Write-* helpers, Format-Elapsed, Stop-RunTranscript
    DmfRequest.ps1  -- Invoke-DmfRequest (REST client with retry)
    DmfAuth.ps1     -- Connect-DmfEnvironment, Get-DmfAuthHeaders, Test-DmfTokenExpiry
    DmfOData.ps1    -- Get-DmfODataAll, New-DmfODataUri
    DmfTemplate.ps1 -- Get-TemplateFolders, Get-TemplateInfo, ConvertTo-DmfTemplateLine
    DmfMetadata.ps1 -- Resolve-DmfEntity, entity-map.json  (OData mode)
    DmfPull.ps1     -- Invoke-DmfEntityPull, Write-DmfEntitySnapshot, _pull.json  (OData mode)

.PARAMETER EnvironmentUrl
    Base URL of the D365 F&O environment (no trailing slash).
    Example: https://contoso.operations.dynamics.com

.PARAMETER TenantId
    Microsoft Entra tenant ID (GUID) or verified domain name.
    Example: contoso.onmicrosoft.com

.PARAMETER LegalEntityId
    D365 legal entity (company) to export from.
    When omitted, the script prompts interactively.
    Example: USMF

.PARAMETER TemplateName
    When supplied, processes this one template without the selection menu.
    The value must match a TemplateId exactly.
    Example: '010 - System Setup'

.PARAMETER All
    Process every validated template in the environment without showing the
    selection menu.  Cannot be combined with -TemplateName.

    Combine with -Force and -LegalEntityId for a fully unattended sweep; when
    -LegalEntityId is omitted the script still prompts for it, so -All -Force
    without -LegalEntityId is rejected up front rather than left to block on a
    prompt that nobody is there to answer.

.PARAMETER DownloadPath
    Directory to save the downloaded export .zip files.
    Defaults to $env:TEMP.  Pass an empty string ('') to skip downloading.

.PARAMETER AuthMode
    How to sign in.  Auto (default): open the default browser when the
    session is interactive, falling back to the device code flow if that
    fails; Browser: browser only; DeviceCode: print a code to enter in
    any browser (for SSH sessions and servers without a browser).

.PARAMETER LogPath
    Path for the run transcript log.  When omitted, a log is auto-generated in
    DownloadPath (or $env:TEMP) as DMFProjectExport_<timestamp>.log.
    Pass an empty string ('') to suppress transcript logging entirely.

.PARAMETER PollIntervalSeconds
    How often (in seconds) to check export status.  Range: 5-300.  Default: 30.

.PARAMETER TimeoutMinutes
    Per-template polling timeout in minutes.  Range: 1-480.  Default: 60.

.PARAMETER MaxRetries
    Maximum number of automatic retries for transient REST failures (HTTP 5xx,
    408, network errors).  HTTP 429 throttling draws on a separate budget --
    see $Script:ThrottleMaxRetries in lib/DmfRequest.ps1.
    Range: 0-10.  Default: 3.

.PARAMETER Force
    Skip the confirmation prompt before exporting.

.PARAMETER WhatIf
    Show what would be created and exported without making any changes to D365.
    When -TemplateName is also provided, no API calls are made at all.
    Otherwise the template list is fetched (read-only) to populate the menu.

.PARAMETER TemplateSource
    Local (default) lists template folders under -ResourcesPath; Environment
    lists the D365 environment's validated templates as before.

.PARAMETER ResourcesPath
    Root scanned for template folders and home of entity-map.json.
    Default: ./resources

.PARAMETER Mode
    Dmf (default): create a DMF project and export a package.
    OData: read the template entities from the OData endpoints into JSON.

.PARAMETER DataPath
    OData mode: root under which <environment>/<legal entity>/ is created.
    Default: ./data  (git-ignored)

.PARAMETER MaxRecordsPerEntity
    OData mode: safety cap per entity (0 = unlimited).  An entity that hits
    the cap is written truncated and reported as Truncated.

.PARAMETER RefreshEntityMap
    OData mode: ignore cached resolutions and re-query the Metadata service
    (manual entries in entity-map.json are still honoured).

.PARAMETER IncludeDisabled
    Include manifest lines marked Disable=true (local templates only).

.PARAMETER PassThru
    Return result objects to the pipeline after completion.
    Dmf mode, one per template: Template, TemplateId, ProjectName, LegalEntityId,
    Status, LinesAdded, ExecutionId, DownloadUrl, DownloadedTo, Elapsed.
    OData mode, one per entity: Templates, Entity, TargetEntity, Collection,
    LegalEntityId, Status, Reason, RecordCount, File, Elapsed.

.EXAMPLE
    # Local template, DMF export (creates the project in D365 from the manifest)
    .\Invoke-ProjectExport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'USMF' `
        -TemplateName   '010 - System Setup' `
        -DownloadPath   'C:\DMF\Downloads'

.EXAMPLE
    # Local template, OData pull into .\data\<env>\USMF\
    .\Invoke-ProjectExport.ps1 `
        -EnvironmentUrl 'https://contoso-uat.sandbox.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'USMF' `
        -TemplateName   '010 - System Setup' `
        -Mode           OData

.EXAMPLE
    # Whole environment via OData: every local template, unattended
    .\Invoke-ProjectExport.ps1 -EnvironmentUrl 'https://contoso.operations.dynamics.com' -TenantId 'contoso.onmicrosoft.com' `
        -LegalEntityId 'USMF' -All -Mode OData -Force -PassThru |
        Where-Object Status -ne 'Pulled' | Format-Table Entity, Status, Reason

.EXAMPLE
    # Preview an OData pull: resolution status per entity from the cache, no sign-in
    .\Invoke-ProjectExport.ps1 -EnvironmentUrl 'https://contoso.operations.dynamics.com' -TenantId 'contoso.onmicrosoft.com' `
        -LegalEntityId 'USMF' -TemplateName '010 - System Setup' -Mode OData -WhatIf

.EXAMPLE
    # Interactive: list templates, prompt for selection and legal entity
    .\Invoke-ProjectExport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com'

.EXAMPLE
    # Supply legal entity up front, choose templates interactively
    .\Invoke-ProjectExport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'USMF'

.EXAMPLE
    # Non-interactive: one template, skip confirm, download to C:\Exports
    .\Invoke-ProjectExport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'USMF' `
        -TemplateName   '010 - System Setup' `
        -DownloadPath   'C:\Exports' `
        -Force

.EXAMPLE
    # WhatIf with a named template -- no API calls at all
    .\Invoke-ProjectExport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'USMF' `
        -TemplateName   '010 - System Setup' `
        -WhatIf

.EXAMPLE
    # Unattended: export every validated template in the environment
    .\Invoke-ProjectExport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'USMF' `
        -DownloadPath   'C:\DMF\Downloads' `
        -All `
        -Force

.EXAMPLE
    # Preview an all-templates sweep -- read-only, nothing created or exported
    .\Invoke-ProjectExport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'USMF' `
        -All `
        -WhatIf
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^https?://')]
    [string]$EnvironmentUrl,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [string]$LegalEntityId,

    [string]$TemplateName,

    [AllowEmptyString()]
    [string]$DownloadPath = $PWD.Path,

    [ValidateSet('Auto', 'Browser', 'DeviceCode')]
    [string]$AuthMode = 'Auto',

    [AllowEmptyString()]
    [string]$LogPath,

    [ValidateRange(5, 300)]
    [int]$PollIntervalSeconds = 30,

    [ValidateRange(1, 480)]
    [int]$TimeoutMinutes = 60,

    [ValidateRange(0, 10)]
    [int]$MaxRetries = 3,

    [switch]$All,

    [ValidateSet('Local', 'Environment')]
    [string]$TemplateSource = 'Local',

    [string]$ResourcesPath,

    [ValidateSet('Dmf', 'OData')]
    [string]$Mode = 'Dmf',

    [string]$DataPath,

    [ValidateRange(0, 100000000)]
    [int]$MaxRecordsPerEntity = 0,

    [switch]$RefreshEntityMap,
    [switch]$IncludeDisabled,

    [switch]$Force,
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
. (Join-Path $libPath 'DmfPull.ps1')

Add-Type -AssemblyName System.IO.Compression.FileSystem

# =============================================================================
#  Pre-flight validation  (before transcript so errors surface cleanly)
# =============================================================================
# The download folder is output: create it rather than demanding it exists.
if ($Mode -eq 'Dmf' -and $DownloadPath -ne '' -and -not (Test-Path $DownloadPath -PathType Container)) {
    if ($WhatIf) { Write-Host "Download path '$DownloadPath' does not exist; it would be created." -ForegroundColor Yellow }
    else         { New-Item -ItemType Directory -Path $DownloadPath -Force | Out-Null }
}

if ($All -and $TemplateName) {
    throw "-All and -TemplateName are mutually exclusive.  Use -All to process every template, or -TemplateName to process exactly one."
}

# Resolved here, not in the param block: on PS 5.1 $PSScriptRoot is empty
# while defaults are evaluated for a script started with a relative path.
if (-not $ResourcesPath) { $ResourcesPath = Join-Path $PSScriptRoot 'resources' }
if (-not $DataPath)      { $DataPath      = Join-Path $PSScriptRoot 'data' }

if ($TemplateSource -eq 'Local' -and -not (Test-Path -LiteralPath $ResourcesPath -PathType Container)) {
    throw "Resources path not found: '$ResourcesPath'.  Capture templates with Export-TemplateDefinition.ps1 or use -TemplateSource Environment."
}
if ($Mode -eq 'OData') {
    foreach ($ignored in 'DownloadPath', 'PollIntervalSeconds', 'TimeoutMinutes') {
        if ($PSBoundParameters.ContainsKey($ignored)) { Write-Host "[WARN] -$ignored is ignored in OData mode." -ForegroundColor Yellow }
    }
}
$entityMapPath = Join-Path $ResourcesPath 'entity-map.json'

# -All -Force signals an unattended run, but the legal entity is prompted for
# when it is not supplied -- which would block forever with nobody watching.
if ($All -and $Force -and -not $LegalEntityId) {
    throw "-All -Force requires -LegalEntityId; without it the script would stop at the legal entity prompt."
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
#  1.  Transcript startup
# =============================================================================
$logBase = if ($DownloadPath -ne '') { $DownloadPath } else { $env:TEMP }
if (-not $PSBoundParameters.ContainsKey('LogPath')) {
    $ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
    $LogPath = Join-Path $logBase "DMFProjectExport_${ts}.log"
}
if ($LogPath -ne '') {
    try {
        Start-Transcript -Path $LogPath -Force | Out-Null
        $Script:TranscriptActive = $true
    } catch { <# transcript not supported in this host -- continue silently #> }
}

# =============================================================================
#  2.  Banner + session info
# =============================================================================
Write-Banner -DryRun:$WhatIf -Title 'D365 F&O Project Export Utility'
Write-Host ''
Write-Info "Environment  : $EnvironmentUrl"
Write-Info "Tenant       : $TenantId"
if ($LegalEntityId) { Write-Info "Legal entity : $LegalEntityId" }
Write-Info "Templates    : $TemplateSource$(if ($TemplateSource -eq 'Local') { "  ($ResourcesPath)" })"
Write-Info "Mode         : $Mode$(if ($Mode -eq 'OData') { '  (read entities via OData; no DMF project is created)' })"
if ($Mode -eq 'OData') {
    Write-Info "Data path    : $DataPath"
} elseif ($DownloadPath -ne '') { Write-Info "Download to  : $DownloadPath" }
else                            { Write-Info 'Download     : disabled' }
if ($Script:TranscriptActive) { Write-Info "Log          : $LogPath" }
if ($WhatIf) { Write-Warn "WhatIf active -- $(if ($Mode -eq 'OData') { 'nothing will be pulled or written' } else { 'no projects will be created or exported' })." }
if ($Force)  { Write-Info 'Force        : confirmation prompt suppressed' }

# =============================================================================
#  3.  WhatIf early exit for named template  (zero API calls)
#      Only needed for the Environment source; a local template list needs no
#      sign-in, so the normal flow already makes zero API calls under -WhatIf.
# =============================================================================
if ($WhatIf -and $TemplateName -and $TemplateSource -eq 'Environment' -and $Mode -eq 'Dmf') {
    $previewLe      = if ($LegalEntityId) { $LegalEntityId } else { '<legal-entity>' }
    $previewProject = "$TemplateName $previewLe"
    Write-Host ''
    Write-Rule 'WhatIf -- no changes will be made'
    Write-Host ''
    Write-Info "Would process template  : $TemplateName"
    Write-Info "Source legal entity     : $previewLe"
    Write-Info "DMF project name        : $previewProject"
    Write-Host ''
    Write-Warn 'Template name is not validated in WhatIf mode; re-run without -WhatIf to confirm it exists.'
    Write-Info 'No API calls made.  Remove -WhatIf to create the project and run the export.'
    if ($PassThru) {
        [pscustomobject]@{
            Template     = $TemplateName
            TemplateId   = $TemplateName
            ProjectName  = $previewProject
            LegalEntityId = $previewLe
            Status       = 'WhatIf'
            LinesAdded   = 0
            ExecutionId  = '-'
            DownloadUrl  = '-'
            DownloadedTo = '-'
            Elapsed      = '-'
        }
    }
    Stop-RunTranscript
    exit 0
}

# =============================================================================
#  4.  Authenticate  (once -- token reused across all templates)
# =============================================================================
$baseUrl     = $EnvironmentUrl.TrimEnd('/')
$dmfBase     = "$baseUrl/data/DataManagementDefinitionGroups/Microsoft.Dynamics.DataEntities"
$session     = $null
$authHeaders = $null

function Connect-IfNeeded {
    # Sign in once, lazily: a local template list needs no session, so with
    # -WhatIf and -TemplateSource Local the script never authenticates.
    if ($null -ne $script:session) { return }
    Write-Step 'Authenticating with Microsoft Entra (device code flow)'
    $script:session = Connect-DmfEnvironment -EnvironmentUrl $baseUrl -TenantId $TenantId -AuthMode $AuthMode
    # Every Invoke-DmfRequest call to this environment reads the session and
    # renews the token silently when it is close to expiry (lib/DmfAuth.ps1).
    $Script:DmfSession  = $script:session
    $script:authHeaders = Get-DmfAuthHeaders -Session $script:session
}

function Get-SelectedEntityLines {
    <#
    .SYNOPSIS  Union of the entity lines of the selected templates (each entity once, first-seen ordering).
    .OUTPUTS   Objects: EntityName, TargetEntity, ExecutionUnit, LevelInExecutionUnit, SequenceInLevel, Templates (List[string])
    #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Templates)

    $union = New-Object System.Collections.Specialized.OrderedDictionary([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($tmpl in $Templates) {
        $lines = @()
        if ($TemplateSource -eq 'Local') {
            $lines = @(ConvertTo-DmfTemplateLine -Lines $tmpl.Lines -IncludeDisabled:$IncludeDisabled)
        } else {
            Connect-IfNeeded
            $uri   = New-DmfODataUri -BaseUrl $baseUrl -Collection 'DefinitionGroupTemplateLines' -Filter "TemplateId eq '$(ConvertTo-DmfODataLiteral $tmpl.TemplateId)'"
            $lines = @(Get-DmfODataAll -Uri $uri -Operation "lines $($tmpl.TemplateId)" -Headers $script:authHeaders)
        }
        foreach ($l in $lines) {
            $name = [string](Get-DmfProp $l 'Entity' '')
            if (-not $name) { continue }
            if ($union.Contains($name)) { $union[$name].Templates.Add([string]$tmpl.TemplateId); continue }
            $entry = [pscustomobject]@{
                EntityName           = $name
                TargetEntity         = Get-DmfProp $l 'TargetEntity'
                ExecutionUnit        = [int](Get-DmfProp $l 'ExecutionUnit' 1)
                LevelInExecutionUnit = [int](Get-DmfProp $l 'LevelInExecutionUnit' 1)
                SequenceInLevel      = [int](Get-DmfProp $l 'Sequence' 1)
                Templates            = [System.Collections.Generic.List[string]]::new()
            }
            $entry.Templates.Add([string]$tmpl.TemplateId)
            $union[$name] = $entry
        }
    }
    return @($union.Values | Sort-Object -Property ExecutionUnit, LevelInExecutionUnit, SequenceInLevel, EntityName)
}

if ($TemplateSource -eq 'Environment') { Connect-IfNeeded }

# =============================================================================
#  5.  Fetch available templates  (OData with pagination)
# =============================================================================
$entityMap = Get-DmfEntityMap -Path $entityMapPath

# The unfiltered local list; -TemplateName resolves against it so a template
# hidden from the menu can still be named explicitly.
$allLocalTemplates = @()

if ($TemplateSource -eq 'Environment') {
    Write-Step 'Fetching available templates'

    $rawTemplates = @(Get-DmfODataAll -Uri "$baseUrl/data/DefinitionGroupTemplateHeaders" -Operation 'list templates' -Headers $authHeaders)

    # Only surface Validated templates; sort by TemplateId client-side.
    $allTemplates = @($rawTemplates |
        Where-Object { (Get-DmfProp $_ 'Status' '') -eq 'Validated' } |
        Sort-Object  -Property TemplateId |
        ForEach-Object -Begin { $i = 1 } -Process {
        [pscustomobject]@{
            Index       = $i++
            TemplateId  = [string]$_.TemplateId
            Description = [string](Get-DmfProp $_ 'Description' '')
            ValidatedOn = $(try { $v = Get-DmfProp $_ 'ValidatedDateTime'; if ($v) { ([datetime]$v).ToString('yyyy-MM-dd') } else { '' } } catch { '' })
            Lines        = $null     # fetched per template when needed
            Folder       = $null
            HasData      = $false
            AppliesTo    = 'Any'     # the environment does not distinguish transports
            SupersededBy = ''
        }
    })

    if ($allTemplates.Count -eq 0) {
        throw 'No validated templates found in this environment.  Ensure DMF definition group templates exist and have Status = Validated.'
    }
    Write-Info "$($allTemplates.Count) validated template(s) found."
}
else {
    Write-Step "Scanning local templates  ($ResourcesPath)"

    $folders      = @(Get-TemplateFolders -ResourcesPath $ResourcesPath)
    $allTemplates = @($(for ($i = 0; $i -lt $folders.Count; $i++) {
        $info = Get-TemplateInfo -Folder $folders[$i] -Index ($i + 1) -EntityMap $entityMap
        foreach ($w in $info.Warnings) { Write-Warn "$($info.Name): $w" }
        $kind = if ($info.HasData) { "$($info.EntityCount) lines, +data" } else { "$($info.EntityCount) lines" }
        if ($Mode -eq 'OData' -and $null -ne $info.ResolvedCount) { $kind += ", $($info.ResolvedCount) resolved" }
        if ($info.IsCustom) { $kind = "custom, $kind" }
        [pscustomobject]@{
            Index          = $info.Index
            TemplateId     = $info.Name
            Description    = $info.Description
            ValidatedOn    = $kind
            Lines          = $info.Lines
            Folder         = $info.Folder.FullName
            HasData        = $info.HasData
            Origin         = $info.Origin
            AppliesTo      = $info.AppliesTo
            SourceTemplate = $info.SourceTemplate
            PullableCount  = $info.PullableCount
            SupersededBy   = ''
            IsValid        = $info.IsValid
        }
    }) | Where-Object { $_.IsValid })

    if ($allTemplates.Count -eq 0) {
        throw "No valid template folders (Manifest.xml) found under '$ResourcesPath'.  Capture some with Export-TemplateDefinition.ps1, or use -TemplateSource Environment."
    }

    # A generated OData companion names the template it came from; record that
    # on the source so it can be pointed at, and hidden, in OData mode.
    foreach ($companion in $allTemplates) {
        if ($companion.AppliesTo -ne 'OData' -or -not $companion.SourceTemplate) { continue }
        $source = $allTemplates | Where-Object { $_.TemplateId -eq $companion.SourceTemplate } | Select-Object -First 1
        if ($null -ne $source) { $source.SupersededBy = $companion.TemplateId }
    }

    # -- Show only the templates that suit the transport --------------------
    # -TemplateName still resolves against the full list, so naming a template
    # explicitly always works.
    $allLocalTemplates = $allTemplates
    $hidden = [System.Collections.Generic.List[string]]::new()
    $allTemplates = @($allTemplates | Where-Object {
        if ($Mode -eq 'OData') {
            # The source of a companion would only report its non-OData
            # entities as skipped, so the companion stands in for it.
            if ($_.SupersededBy) { $hidden.Add($_.TemplateId); return $false }
            # Nothing in it can be pulled (the map is certain, not merely silent).
            if ($null -ne $_.PullableCount -and $_.PullableCount -eq 0) { $hidden.Add($_.TemplateId); return $false }
            return $true
        }
        # Dmf: a companion is a subset built for the other transport.
        if ($_.AppliesTo -eq 'OData') { $hidden.Add($_.TemplateId); return $false }
        return $true
    })

    if ($allTemplates.Count -eq 0) {
        throw "No template under '$ResourcesPath' applies to -Mode $Mode.  $($hidden.Count) template(s) were hidden as belonging to the other transport; name one explicitly with -TemplateName to use it anyway."
    }
    # Re-index after filtering so menu numbers are contiguous.
    $n = 1; foreach ($t in $allTemplates) { $t.Index = $n++ }
    Write-Info "$($allTemplates.Count) local template(s) apply to -Mode $Mode."
    if ($hidden.Count -gt 0) {
        $other = if ($Mode -eq 'OData') { 'superseded by an OData companion, or with no OData-readable entity' } else { 'OData companions' }
        Write-Detail "$($hidden.Count) hidden ($other); name one with -TemplateName to use it anyway."
    }
}

# =============================================================================
#  6.  Template selection
# =============================================================================
$selectedTemplates = [System.Collections.Generic.List[pscustomobject]]::new()

if ($TemplateName) {
    # ── Single-template mode (non-interactive) ──────────────────────────────
    # Resolve against every local template, including the ones the mode filter
    # hides, so an explicit name is always honoured.
    $lookupPool = $(if ($allLocalTemplates.Count -gt 0) { $allLocalTemplates } else { $allTemplates })
    $match = $lookupPool | Where-Object { $_.TemplateId -eq $TemplateName }
    if (-not $match) {
        $available = ($lookupPool | ForEach-Object { "    '$($_.TemplateId)'" }) -join [System.Environment]::NewLine
        $where     = if ($TemplateSource -eq 'Local') { "under '$ResourcesPath'" } else { "in environment '$EnvironmentUrl'" }
        throw "Template '$TemplateName' not found $where.`nAvailable templates:`n$available"
    }
    $selectedTemplates.Add($match)
    Write-Info "Template : $TemplateName"

    # Point at the better-suited template rather than silently doing less.
    if ($Mode -eq 'OData' -and $match.PSObject.Properties['SupersededBy'] -and $match.SupersededBy) {
        Write-Info "'$($match.SupersededBy)' is the OData companion for this template; it drops the entities that cannot be pulled."
    }
    if ($Mode -eq 'Dmf' -and $match.PSObject.Properties['AppliesTo'] -and $match.AppliesTo -eq 'OData') {
        Write-Warn "'$TemplateName' is an OData companion (a subset of '$($match.SourceTemplate)'); a DMF export of it will be missing the entities that were dropped."
    }
} elseif ($All) {
    # ── All-templates mode (non-interactive) ────────────────────────────────
    foreach ($tmpl in $allTemplates) { $selectedTemplates.Add($tmpl) }
    Write-Info "All $($allTemplates.Count) validated template(s) selected (-All)."
} else {
    # ── Interactive numbered selection ─────────────────────────────────────
    Write-Step 'Select templates to export'

    $idWidth   = ($allTemplates | ForEach-Object { $_.TemplateId.Length }   | Measure-Object -Maximum).Maximum
    $idWidth   = [int][Math]::Max(20, [Math]::Min($idWidth, [int]([Math]::Floor(($Script:LineWidth - 22) * 0.45))))
    $descWidth = [int][Math]::Max(20, [Math]::Min(
        ($allTemplates | ForEach-Object { $_.Description.Length } | Measure-Object -Maximum).Maximum,
        $Script:LineWidth - $idWidth - 22
    ))

    $colFmt  = '  {0,3}  {1}  {2}  {3}'
    $lastHdr = if ($TemplateSource -eq 'Local') { 'Contents' } else { 'Validated' }
    Write-Host ''
    Write-Host ($colFmt -f '#', 'Template ID'.PadRight($idWidth), 'Description'.PadRight($descWidth), $lastHdr) -ForegroundColor White
    Write-Host ($colFmt -f '---', ('-' * $idWidth), ('-' * $descWidth), '----------') -ForegroundColor DarkGray

    foreach ($tmpl in $allTemplates) {
        $idCol   = if ($tmpl.TemplateId.Length -gt $idWidth) {
                       $tmpl.TemplateId.Substring(0, $idWidth - 3) + '...'
                   } else { $tmpl.TemplateId.PadRight($idWidth) }
        $descCol = if ($tmpl.Description.Length -gt $descWidth) {
                       $tmpl.Description.Substring(0, $descWidth - 3) + '...'
                   } else { $tmpl.Description.PadRight($descWidth) }
        Write-Host ($colFmt -f $tmpl.Index, $idCol, $descCol, $tmpl.ValidatedOn) -ForegroundColor White
    }

    # Validated input loop
    $selectedIndices = $null
    do {
        Write-Host ''
        $rawInput = (Read-Host '  Selection  (e.g. 1   1,3   2-4   A=all   Q=quit)').Trim().ToLower()

        if ($rawInput -in 'q', 'quit') { $selectedIndices = @(); break }
        if ($rawInput -in 'a', 'all', '') { $selectedIndices = 1..$allTemplates.Count; break }

        $parsed   = [System.Collections.Generic.List[int]]::new()
        $badInput = $false

        foreach ($token in ($rawInput -split ',')) {
            $token = $token.Trim()
            if ($token -match '^(\d+)\s*-\s*(\d+)$') {
                $start = [int]$Matches[1]; $end = [int]$Matches[2]
                if ($start -gt $end) {
                    Write-Warn "  Invalid range '${start}-${end}': start must be <= end."
                    $badInput = $true; break
                }
                $start..$end | ForEach-Object { $parsed.Add($_) }
            } elseif ($token -match '^\d+$') {
                $parsed.Add([int]$token)
            } else {
                Write-Warn "  Unrecognised input '$token'.  Use numbers, ranges (1-3), commas, A, or Q."
                $badInput = $true; break
            }
        }

        if ($badInput) { continue }

        $outOfRange = @($parsed | Where-Object { $_ -lt 1 -or $_ -gt $allTemplates.Count })
        if ($outOfRange.Count -gt 0) {
            Write-Warn "  Out-of-range: $($outOfRange -join ', ').  Valid: 1-$($allTemplates.Count)."
            continue
        }
        if ($parsed.Count -eq 0) {
            Write-Warn '  Nothing selected.  Enter numbers, A for all, or Q to quit.'
            continue
        }

        $selectedIndices = @($parsed | Sort-Object -Unique)
        break
    } while ($true)

    if (-not $selectedIndices -or $selectedIndices.Count -eq 0) {
        Write-Host ''
        Write-Warn 'No templates selected.  Exiting.'
        Stop-RunTranscript
        exit 0
    }

    foreach ($idx in $selectedIndices) {
        $selectedTemplates.Add(($allTemplates | Where-Object { $_.Index -eq $idx }))
    }
}

# =============================================================================
#  7.  Legal entity prompt  (when not supplied as a parameter)
# =============================================================================
if (-not $LegalEntityId) {
    Write-Host ''
    do {
        $LegalEntityId = (Read-Host '  Export from legal entity (e.g. USMF)').Trim()
    } while (-not $LegalEntityId)
}

# =============================================================================
#  8.  WhatIf exit  (interactive path -- template list was already fetched)
# =============================================================================
if ($WhatIf -and $Mode -eq 'OData') {
    $envName = Get-DmfEnvironmentName -EnvironmentUrl $baseUrl
    $folder  = Get-DmfPullFolder -DataPath $DataPath -EnvironmentName $envName -LegalEntityId $LegalEntityId
    $lines   = @(Get-SelectedEntityLines -Templates $selectedTemplates)
    $preview = @(Resolve-DmfEntity -Lines $lines -Map $entityMap -Offline)

    Write-Host ''
    Write-Rule 'WhatIf -- nothing will be pulled or written'
    Write-Host ''
    Write-Info "Environment  : $EnvironmentUrl  ($envName)"
    Write-Info "Legal entity : $LegalEntityId"
    Write-Info "Templates    : $(($selectedTemplates | ForEach-Object { $_.TemplateId }) -join ', ')"
    Write-Info "Data folder  : $folder"
    Write-Info "Entities     : $($lines.Count)  (resolution shown from entity-map.json only)"
    Write-Host ''
    foreach ($p in $preview) {
        $file   = Join-Path $folder (Get-DmfSnapshotFileName -EntityName $p.EntityName)
        $exists = if (Test-Path -LiteralPath $file -PathType Leaf) { 'would overwrite' } else { 'new file' }
        $how    = switch ($p.Status) {
            'Resolved'  { "-> $($p.Collection)$(if ($p.CompanySpecific) { ' (per company)' })" }
            'NotPublic' { 'not OData-enabled -- would be skipped' }
            default     { 'would resolve via the Metadata service' }
        }
        Write-Detail ("{0,-45} {1,-48} {2}" -f $p.EntityName, $how, $exists)
    }
    Write-Host ''
    $apiNote = if ($TemplateSource -eq 'Environment') { 'Only the template list and lines were read.' } else { 'No API calls made.' }
    Write-Info "$apiNote  Remove -WhatIf to pull."
    if ($PassThru) {
        foreach ($p in $preview) {
            [pscustomobject]@{
                Templates     = ''
                Entity        = $p.EntityName
                TargetEntity  = $p.TargetEntity
                Collection    = $p.Collection
                LegalEntityId = $LegalEntityId
                Status        = 'WhatIf'
                Reason        = $p.Status
                RecordCount   = 0
                File          = Join-Path $folder (Get-DmfSnapshotFileName -EntityName $p.EntityName)
                Elapsed       = '-'
            }
        }
    }
    Stop-RunTranscript
    exit 0
}

if ($WhatIf) {
    Write-Host ''
    Write-Rule 'WhatIf -- no changes will be made'
    Write-Host ''
    Write-Info "Environment  : $EnvironmentUrl"
    Write-Info "Legal entity : $LegalEntityId"
    Write-Info "Templates    : $($selectedTemplates.Count)"
    foreach ($tmpl in $selectedTemplates) {
        $projectName = "$($tmpl.TemplateId) $LegalEntityId"
        $meta = "project: $projectName"
        if ($tmpl.Description) { $meta += "  -- $($tmpl.Description)" }
        if ($tmpl.ValidatedOn)  { $meta += $(if ($TemplateSource -eq 'Local') { "  [$($tmpl.ValidatedOn)]" } else { "  [validated: $($tmpl.ValidatedOn)]" }) }
        Write-Detail "[$($tmpl.Index)] $meta"
    }
    Write-Host ''
    Write-Info 'No projects created or export jobs submitted.  Remove -WhatIf to proceed.'
    if ($PassThru) {
        $selectedTemplates | ForEach-Object {
            [pscustomobject]@{
                Template      = $_.Description
                TemplateId    = $_.TemplateId
                ProjectName   = "$($_.TemplateId) $LegalEntityId"
                LegalEntityId = $LegalEntityId
                Status        = 'WhatIf'
                LinesAdded    = 0
                ExecutionId   = '-'
                DownloadUrl   = '-'
                DownloadedTo  = '-'
                Elapsed       = '-'
            }
        }
    }
    Stop-RunTranscript
    exit 0
}

# =============================================================================
#  9.  Confirmation  (skipped with -Force)
# =============================================================================
Write-Host ''
Write-Rule $(if ($Mode -eq 'OData') { 'Ready to pull (OData)' } else { 'Ready to export' })
Write-Host ''
Write-Info "Environment  : $EnvironmentUrl"
Write-Info "Legal entity : $LegalEntityId"
Write-Info "Templates    : $($selectedTemplates.Count)"
if ($Mode -eq 'OData') {
    Write-Info "Data folder  : $(Get-DmfPullFolder -DataPath $DataPath -EnvironmentName (Get-DmfEnvironmentName -EnvironmentUrl $baseUrl) -LegalEntityId $LegalEntityId)"
}

foreach ($tmpl in $selectedTemplates) {
    $meta = if ($Mode -eq 'OData') { $tmpl.TemplateId } else { "$($tmpl.TemplateId) $LegalEntityId" }
    if ($tmpl.Description) { $meta += "  -- $($tmpl.Description)" }
    if ($tmpl.ValidatedOn) { $meta += $(if ($TemplateSource -eq 'Local') { "  [$($tmpl.ValidatedOn)]" } else { "  [validated: $($tmpl.ValidatedOn)]" }) }
    Write-Detail "[$($tmpl.Index)] $meta"
}

if (-not $Force) {
    Write-Host ''
    $confirm = (Read-Host '  Proceed? [Y]es / [N]o  (default: Y)').Trim().ToUpper()
    if ($confirm -in 'N', 'NO') {
        Write-Info 'Cancelled.'
        Stop-RunTranscript
        exit 0
    }
}

Connect-IfNeeded

# =============================================================================
#  10a.  OData pull  (-Mode OData)  -- then exit
# =============================================================================
if ($Mode -eq 'OData') {
    $scriptStart   = Get-Date
    $envName       = $session.EnvironmentName
    $legalEntity   = $LegalEntityId.ToUpperInvariant()
    $folder        = Get-DmfPullFolder -DataPath $DataPath -EnvironmentName $envName -LegalEntityId $legalEntity -Create
    $templateNames = @($selectedTemplates | ForEach-Object { [string]$_.TemplateId })

    $lines = @(Get-SelectedEntityLines -Templates $selectedTemplates)
    $lineByName = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($l in $lines) { $lineByName[$l.EntityName] = $l }

    # ── Resolve every entity once, up front ─────────────────────────────────
    Write-Step "Resolving $($lines.Count) entities  (entity map, then Metadata service)"
    $resolutions = @(Resolve-DmfEntity -Lines $lines -Session $session -Map $entityMap -Refresh:$RefreshEntityMap)
    [void](Save-DmfEntityMap -Map $entityMap)

    $nResolved  = @($resolutions | Where-Object { $_.Status -eq 'Resolved' }).Count
    $nNotPublic = @($resolutions | Where-Object { $_.Status -eq 'NotPublic' }).Count
    $unresolved = @($resolutions | Where-Object { $_.Status -like 'Unresolved*' })
    Write-Info "$nResolved resolved  |  $nNotPublic not OData-enabled  |  $($unresolved.Count) unresolved"
    Write-Detail "Entity map saved: $entityMapPath  ($($entityMap.Entities.Count) entries)"
    foreach ($u in $unresolved) { Write-Warn "  unresolved: $($u.EntityName) -- $($u.Reason)" }
    if ($nResolved -eq 0) {
        throw 'No entity could be resolved to an OData collection; nothing to pull.  Check entity-map.json and the Metadata service (Invoke-EnvironmentProbe.ps1).'
    }
    if (-not $Force -and ($unresolved.Count -gt 0 -or $nNotPublic -gt 0)) {
        Write-Host ''
        $confirm = (Read-Host "  Pull the $nResolved resolved entities and skip the rest? [Y]es / [N]o  (default: Y)").Trim().ToUpper()
        if ($confirm -in 'N', 'NO') { Write-Info 'Cancelled.'; Stop-RunTranscript; exit 0 }
    }

    # ── Pull ────────────────────────────────────────────────────────────────
    $index = Read-DmfPullIndex -Folder $folder
    $index.environment    = $envName
    $index.environmentUrl = $baseUrl
    $index.legalEntity    = $legalEntity
    $runStartedAt = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')

    $pullResults = [System.Collections.Generic.List[pscustomobject]]::new()
    $i = 0
    foreach ($r in $resolutions) {
        $i++
        $tpls = @($lineByName[$r.EntityName].Templates)
        $pr = [pscustomobject]@{
            Templates     = ($tpls -join '; ')
            Entity        = $r.EntityName
            TargetEntity  = $r.TargetEntity
            Collection    = $r.Collection
            LegalEntityId = $legalEntity
            Status        = 'Skipped'
            Reason        = $r.Reason
            RecordCount   = 0
            File          = '-'
            Elapsed       = '-'
        }

        if ($r.Status -ne 'Resolved') {
            $pr.Status = if ($r.Status -eq 'NotPublic') { 'NotPublic' } else { 'Unresolved' }
            Set-DmfPullIndexEntry -Index $index -EntityName $r.EntityName -Status $pr.Status -Reason $r.Reason -Templates $tpls
            Write-Host ("  [{0}/{1}]  {2,-45} {3}" -f $i, $resolutions.Count, $r.EntityName, "$($pr.Status) -- $($r.Reason)") -ForegroundColor DarkGray
            $pullResults.Add($pr)
            if ($PassThru) { Write-Output $pr }
            continue
        }

        Write-Step ("[{0}/{1}]  {2}  ({3})" -f $i, $resolutions.Count, $r.EntityName, $r.Collection)
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $pull = Invoke-DmfEntityPull -Session $session -Resolution $r -LegalEntityId $legalEntity -MaxRecords $MaxRecordsPerEntity
            $file = Write-DmfEntitySnapshot -Folder $folder -Resolution $r -Session $session -LegalEntityId $legalEntity -Records $pull.Records -ElapsedSeconds $sw.Elapsed.TotalSeconds -Truncated $pull.Truncated
            $sw.Stop()
            $pr.Status      = if ($pull.Truncated) { 'Truncated' } else { 'Pulled' }
            $pr.Reason      = if ($pull.Truncated) { "capped at $MaxRecordsPerEntity records" } else { '' }
            $pr.RecordCount = $pull.Records.Count
            $pr.File        = $file
            Set-DmfPullIndexEntry -Index $index -EntityName $r.EntityName -Status $pr.Status -Reason $pr.Reason -File (Split-Path -Leaf $file) -RecordCount $pr.RecordCount -Templates $tpls
            $msg = "{0:N0} record(s) in {1}  ->  {2}" -f $pr.RecordCount, (Format-Elapsed $sw.Elapsed), (Split-Path -Leaf $file)
            if ($pull.Truncated) { Write-Warn "$msg  (TRUNCATED)" } else { Write-Info $msg }
        }
        catch {
            $sw.Stop()
            $pr.Status = 'Failed'
            $pr.Reason = $_.Exception.Message
            Set-DmfPullIndexEntry -Index $index -EntityName $r.EntityName -Status 'Failed' -Reason $pr.Reason -Templates $tpls
            Write-Fail "  $($r.EntityName): $($_.Exception.Message)"
            Write-Verbose $_.ScriptStackTrace
        }
        $pr.Elapsed = Format-Elapsed $sw.Elapsed
        $pullResults.Add($pr)
        if ($PassThru) { Write-Output $pr }
    }

    $index.lastRun = [ordered]@{
        startedAt  = $runStartedAt
        finishedAt = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        templates  = $templateNames
        tool       = "Invoke-ProjectExport.ps1 $($Script:Version)"
    }
    Write-DmfPullIndex -Folder $folder -Index $index

    # ── Summary ─────────────────────────────────────────────────────────────
    $totalElapsed = Format-Elapsed ((Get-Date) - $scriptStart)
    $divider      = '=' * $Script:LineWidth
    Write-Host ''
    Write-Host $divider -ForegroundColor DarkCyan
    Write-Host "  Summary  --  $($pullResults.Count) entity/entities  |  $folder  |  total time: $totalElapsed" -ForegroundColor Cyan
    Write-Host $divider -ForegroundColor DarkCyan
    Write-Host ''
    Write-Host ('  {0,-12} {1,10}  {2,-8}  {3}' -f 'Status', 'Records', 'Elapsed', 'Entity') -ForegroundColor White
    Write-Host ('  {0,-12} {1,10}  {2,-8}  {3}' -f ('-' * 12), ('-' * 10), ('-' * 8), ('-' * 40)) -ForegroundColor DarkGray
    $nFailed = 0
    foreach ($pr in $pullResults) {
        $colour = switch ($pr.Status) { 'Pulled' { 'Green' } 'Truncated' { 'Yellow' } 'Failed' { 'Red' } default { 'DarkGray' } }
        if ($pr.Status -eq 'Failed') { $nFailed++ }
        $recs = if ($pr.Status -in 'Pulled', 'Truncated') { '{0:N0}' -f $pr.RecordCount } else { '-' }
        $note = if ($pr.Reason -and $pr.Status -notin 'Pulled') { "  -- $($pr.Reason)" } else { '' }
        Write-Host ('  {0,-12} {1,10}  {2,-8}  {3}{4}' -f $pr.Status, $recs, $pr.Elapsed, $pr.Entity, $note) -ForegroundColor $colour
    }
    $nPulled = @($pullResults | Where-Object { $_.Status -in 'Pulled', 'Truncated' }).Count
    Write-Host ''
    Write-Host $divider -ForegroundColor DarkCyan
    Write-Host "  $nPulled pulled  |  $nNotPublic not OData-enabled  |  $($unresolved.Count) unresolved  |  $nFailed failed" -ForegroundColor $(if ($nFailed -eq 0) { 'Green' } else { 'Red' })
    Write-Host "  Run index: $(Join-Path $folder '_pull.json')" -ForegroundColor Gray
    Write-Host $divider -ForegroundColor DarkCyan
    Write-Host ''

    Stop-RunTranscript
    if ($nFailed -gt 0) { exit 1 }
    exit 0
}

# =============================================================================
#  10.  Process templates  (-Mode Dmf)
# =============================================================================
$results     = [System.Collections.Generic.List[pscustomobject]]::new()
$scriptStart = Get-Date

foreach ($tmpl in $selectedTemplates) {
    $tmplId      = $tmpl.TemplateId
    $projectName = "$tmplId $LegalEntityId"
    $tmplSafe    = ($projectName -replace '[^A-Za-z0-9]', '-') -replace '-{2,}', '-'
    $tmplStart   = Get-Date
    $timestamp   = Get-Date -Format 'yyyyMMddHHmmss'
    $tmplResult  = [pscustomobject]@{
        Template      = $tmpl.Description
        TemplateId    = $tmplId
        ProjectName   = $projectName
        LegalEntityId = $LegalEntityId
        Status        = 'Error'
        LinesAdded    = 0
        ExecutionId   = '-'
        DownloadUrl   = '-'
        DownloadedTo  = '-'
        Elapsed       = '-'
    }

    $divider = '=' * $Script:LineWidth
    Write-Host ''
    Write-Host $divider -ForegroundColor DarkCyan
    Write-Host "  [$($results.Count + 1)/$($selectedTemplates.Count)]  $tmplId" -ForegroundColor Cyan
    if ($tmpl.Description) {
        Write-Host "  $($tmpl.Description)" -ForegroundColor DarkGray
    }
    Write-Host "  Project: $projectName" -ForegroundColor DarkGray
    Write-Host $divider -ForegroundColor DarkCyan

    # Token expiry: renewed silently when a refresh token is available,
    # otherwise warned about as before.
    [void](Test-DmfTokenExpiry -Session $session -Activity 'export')

    try {
        # ── a. Read template lines  (local Manifest.xml or DefinitionGroupTemplateLines) ──
        $rawLines = [System.Collections.Generic.List[psobject]]::new()
        if ($TemplateSource -eq 'Local') {
            Write-Step "Reading manifest  ($tmplId)"
            foreach ($ln in (ConvertTo-DmfTemplateLine -Lines $tmpl.Lines -IncludeDisabled:$IncludeDisabled)) { $rawLines.Add($ln) }
            $disabledCount = @($tmpl.Lines | Where-Object { $_.Disable }).Count
            if ($disabledCount -gt 0 -and -not $IncludeDisabled) {
                Write-Info "$disabledCount disabled line(s) skipped (use -IncludeDisabled to keep them)."
            }
        } else {
            Write-Step "Fetching template lines  ($tmplId)"
            $linesUrl = New-DmfODataUri -BaseUrl $baseUrl -Collection 'DefinitionGroupTemplateLines' -Filter "TemplateId eq '$(ConvertTo-DmfODataLiteral $tmplId)'"
            foreach ($ln in (Get-DmfODataAll -Uri $linesUrl -Operation 'list template lines' -Headers $authHeaders)) { $rawLines.Add($ln) }
        }

        if ($rawLines.Count -eq 0) {
            Write-Warn "No template lines found for '$tmplId' -- skipping."
            $tmplResult.Status = 'Skipped'
            $results.Add($tmplResult)
            if ($PassThru) { Write-Output $tmplResult }
            continue
        }

        Write-Info "$($rawLines.Count) line(s) found."

        # ── b. Delete existing project  (check first; D365 returns 400 on missing record) ───
        Write-Step "Removing existing project  (if any)"

        $encodedName    = [System.Uri]::EscapeDataString($projectName)
        $projectExists  = $false
        try {
            Invoke-DmfRequest -Operation 'CheckProject' -Params @{
                Method  = 'Get'
                Uri     = "$baseUrl/data/DataManagementDefinitionGroups('$encodedName')"
                Headers = $authHeaders
            } | Out-Null
            $projectExists = $true
        }
        catch {
            # 404 = standard not-found; 400 = D365 also uses this when record is missing
            if ($_ -match 'HTTP 404' -or $_ -match 'HTTP 400') {
                Write-Info "No existing project -- will create fresh."
            } else {
                throw
            }
        }

        if ($projectExists) {
            # Piped to Out-Null: a DELETE returns an empty body, and without this
            # that empty string would be the first object on the -PassThru pipeline.
            Invoke-DmfRequest -Operation 'DeleteProject' -Params @{
                Method  = 'Delete'
                Uri     = "$baseUrl/data/DataManagementDefinitionGroups('$encodedName')"
                Headers = $authHeaders
            } | Out-Null
            Write-Info "Deleted existing project."
        }

        # ── c. Create new export project ──────────────────────────────────────
        Write-Step "Creating export project  '$projectName'"

        $description = if ($tmpl.PSObject.Properties['Description'] -and $tmpl.Description) { $tmpl.Description } else { '' }

        Invoke-DmfRequest -Operation 'CreateProject' -Params @{
            Method      = 'Post'
            Uri         = "$baseUrl/data/DataManagementDefinitionGroups"
            Headers     = $authHeaders
            ContentType = 'application/json'
            Body        = (@{
                Name                = $projectName
                ProjectCategory     = 'Project'
                OperationType       = 'Export'
                GenerateDataPackage = 'No'
                Description         = $description
                TruncateEntityData  = 'No'
            } | ConvertTo-Json)
        } | Out-Null

        Write-Info "Project created."

        # ── d. Add entity records (one per template line) ─────────────────────
        Write-Step "Adding entities  ($($rawLines.Count) lines)"

        $addedCount = 0
        $addErrors  = 0

        foreach ($line in $rawLines) {
            $entityName = $line.Entity
            $lineNum    = $addedCount + $addErrors + 1
            Write-Detail "  [$lineNum/$($rawLines.Count)]  $entityName"

            $detailBody = @{
                DefinitionGroupId        = $projectName
                EntityName               = $entityName
                ExecutionUnit            = $line.ExecutionUnit
                LevelInExecutionUnit     = $line.LevelInExecutionUnit
                SequenceInLevel          = $line.Sequence
                FailLevelOnError         = $line.FailLevelOnError
                FailExecutionUnitOnError = $line.FailExecutionUnitOnError
                RunValidateField         = 'Yes'
                RunBusinessValidation    = 'Yes'
                RunBusinessLogic         = 'Yes'
                SkipStaging              = 'Yes'
                IsTransformed            = 'No'
                DefaultRefreshType       = 'FullPush'
                Disable                  = 'No'
                AutoGenerateMapping      = 'Yes'
                SourceFormat             = 'EXCEL'
            }

            $entityAdded   = $false
            $skipStaging   = 'Yes'
            $conflictTries = 0
            $maxConflict   = 3

            while (-not $entityAdded) {
                $detailBody['SkipStaging'] = $skipStaging
                try {
                    Invoke-DmfRequest -Operation "AddEntity ($entityName)" -Params @{
                        Method      = 'Post'
                        Uri         = "$baseUrl/data/DataManagementDefinitionGroupDetails"
                        Headers     = $authHeaders
                        ContentType = 'application/json'
                        Body        = ($detailBody | ConvertTo-Json)
                    } | Out-Null
                    $addedCount++
                    $entityAdded = $true
                }
                catch {
                    if ($_ -match 'Staging cannot be skipped' -and $skipStaging -eq 'Yes') {
                        Write-Detail "  [$lineNum/$($rawLines.Count)]  Staging not supported -- retrying with SkipStaging=No"
                        $skipStaging = 'No'
                    }
                    elseif ($_ -match 'update conflict' -and $conflictTries -lt $maxConflict) {
                        $conflictTries++
                        Write-Detail "  [$lineNum/$($rawLines.Count)]  Update conflict -- retrying in 5s ($conflictTries/$maxConflict)"
                        Start-Sleep -Seconds 5
                    }
                    else {
                        $addErrors++
                        Write-Warn "  Failed to add entity '$entityName': $_"
                        $entityAdded = $true
                    }
                }
            }
        }

        $tmplResult.LinesAdded = $addedCount

        if ($addedCount -eq 0) {
            throw "No entities were added to project '$projectName' -- cannot export an empty project."
        }
        if ($addErrors -gt 0) {
            Write-Warn "$addErrors entity addition(s) failed; $addedCount entity/entities added successfully."
        } else {
            Write-OK "All $addedCount entities added."
        }

        # ── e. Submit export job ──────────────────────────────────────────────
        $packageName = "${tmplSafe}_${timestamp}"
        Write-Step "Submitting export job  (company: $LegalEntityId)"
        Write-Info "Package name : $packageName"

        $exportResp = Invoke-DmfRequest -Operation 'ExportToPackage' -Params @{
            Method      = 'Post'
            Uri         = "$dmfBase.ExportToPackage"
            Headers     = $authHeaders
            ContentType = 'application/json'
            Body        = (@{
                definitionGroupId = $projectName
                packageName       = $packageName
                executionId       = ''
                reExecute         = $false
                legalEntityId     = $LegalEntityId
            } | ConvertTo-Json)
        }

        $executionId = $exportResp.value
        if ([string]::IsNullOrWhiteSpace($executionId)) {
            throw "ExportToPackage returned an empty execution ID.  Response: $($exportResp | ConvertTo-Json -Depth 5)"
        }
        Write-Info "Execution ID : $executionId"
        $tmplResult.ExecutionId = $executionId

        # ── f. Poll for completion ────────────────────────────────────────────
        Write-Step "Polling for completion  (every ${PollIntervalSeconds}s  |  timeout ${TimeoutMinutes}min)"

        $terminalStatuses = @('Succeeded', 'PartiallySucceeded', 'Failed', 'Canceled')
        $statusBody       = @{ executionId = $executionId } | ConvertTo-Json
        $deadline         = (Get-Date).AddMinutes($TimeoutMinutes)
        $pollStart        = Get-Date
        $lastStatus       = ''
        $spinChars        = @('|', '/', '-', '\')
        $spinIdx          = 0

        try {
            while ((Get-Date) -lt $deadline) {
                Start-Sleep -Seconds $PollIntervalSeconds

                $statusResp = Invoke-DmfRequest -Operation 'GetExecutionSummaryStatus' -Params @{
                    Method      = 'Post'
                    Uri         = "$dmfBase.GetExecutionSummaryStatus"
                    Headers     = $authHeaders
                    ContentType = 'application/json'
                    Body        = $statusBody
                }

                $currentStatus = $statusResp.value
                $elapsed       = (Get-Date) - $pollStart
                $timeoutSecs   = $TimeoutMinutes * 60
                $remaining     = [Math]::Max(0, $timeoutSecs - $elapsed.TotalSeconds)
                $pct           = [Math]::Min(99, ($elapsed.TotalSeconds / $timeoutSecs) * 100)
                $spinChar      = $spinChars[$spinIdx % $spinChars.Count]
                $spinIdx++

                Write-Progress `
                    -Id          1 `
                    -Activity    "Exporting: $tmplId" `
                    -Status      "$spinChar  $currentStatus  |  Elapsed: $(Format-Elapsed $elapsed)  |  Timeout in: $(Format-Elapsed ([TimeSpan]::FromSeconds($remaining)))" `
                    -PercentComplete $pct

                if ($currentStatus -ne $lastStatus) {
                    Write-Info "  [$(Get-Date -Format 'HH:mm:ss')]  Status: $currentStatus"
                    $lastStatus = $currentStatus
                }

                if ($terminalStatuses -contains $currentStatus) { break }
            }
        }
        finally {
            Write-Progress -Id 1 -Activity "Exporting: $tmplId" -Completed
        }

        $tmplResult.Status = if ($lastStatus) { $lastStatus } else { 'TimedOut' }

        # ── g. Retrieve download URL  (on success only) ───────────────────────
        if ($lastStatus -in 'Succeeded', 'PartiallySucceeded') {
            Write-Step 'Retrieving download URL'

            $urlResp = Invoke-DmfRequest -Operation 'GetExportedPackageUrl' -Params @{
                Method      = 'Post'
                Uri         = "$dmfBase.GetExportedPackageUrl"
                Headers     = $authHeaders
                ContentType = 'application/json'
                Body        = (@{ executionId = $executionId } | ConvertTo-Json)
            }

            $downloadUrl = $urlResp.value
            if (-not [string]::IsNullOrWhiteSpace($downloadUrl)) {
                $tmplResult.DownloadUrl = $downloadUrl
                Write-Info 'Download URL retrieved.'
                Write-Detail $downloadUrl

                # ── h. Download and extract package ──────────────────────────
                if ($DownloadPath -ne '') {
                    Write-Step 'Downloading and extracting package'
                    $zipFile     = Join-Path $DownloadPath "${tmplSafe}_${timestamp}.zip"
                    $extractPath = Join-Path $DownloadPath $packageName
                    try {
                        # Routed through Invoke-DmfDownload so a throttled or
                        # flaky blob endpoint backs off and resumes instead of
                        # failing the whole export.
                        Invoke-DmfDownload -Uri $downloadUrl -OutFile $zipFile -Operation 'download package'
                        $fileSizeMB = [Math]::Round((Get-Item $zipFile).Length / 1MB, 2)
                        Write-Info "Downloaded  : $zipFile  ($fileSizeMB MB)"

                        # Extract into $DownloadPath\$packageName
                        if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
                        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipFile, $extractPath)
                        Write-OK "Extracted   : $extractPath"
                        $tmplResult.DownloadedTo = $extractPath

                        # Remove zip now that extraction succeeded
                        Remove-Item $zipFile -Force
                    } catch {
                        Write-Warn "Download/extract failed: $_"
                        Write-Info 'Use the download URL above to retrieve the package manually.'
                    }
                }
            } else {
                Write-Warn 'GetExportedPackageUrl returned an empty URL.'
            }
        }

        # ── Outcome message ───────────────────────────────────────────────────
        $tmplElapsed        = (Get-Date) - $tmplStart
        $tmplResult.Elapsed = Format-Elapsed $tmplElapsed
        $jobPath            = "Data management > Job history > execution ID: $executionId"

        switch ($lastStatus) {
            'Succeeded' {
                Write-OK "Export completed successfully.  ($($tmplResult.Elapsed))"
            }
            'PartiallySucceeded' {
                Write-Warn 'Export partially succeeded -- some data may have been skipped or errored.'
                Write-Info "Review : $jobPath"
            }
            'Failed' {
                Write-Fail 'Export failed.'
                Write-Info "Review : $jobPath"
            }
            'Canceled' {
                Write-Fail 'Export was canceled.'
                Write-Info "Execution ID : $executionId"
            }
            default {
                $tmplResult.Status = 'TimedOut'
                Write-Warn "Timed out after $TimeoutMinutes minutes.  Last known status: '$lastStatus'"
                Write-Info "Execution ID : $executionId"
            }
        }
    }
    catch {
        $tmplResult.Elapsed = Format-Elapsed ((Get-Date) - $tmplStart)
        $tmplResult.Status  = 'Error'
        Write-Fail "Template '$tmplId' failed:"
        Write-Fail "  $_"
        Write-Verbose $_.ScriptStackTrace
    }

    $results.Add($tmplResult)
    if ($PassThru) { Write-Output $tmplResult }
}

# =============================================================================
#  11.  Summary
# =============================================================================
$totalElapsed = Format-Elapsed ((Get-Date) - $scriptStart)
$divider      = '=' * $Script:LineWidth

Write-Host ''
Write-Host $divider -ForegroundColor DarkCyan
Write-Host "  Summary  --  $($results.Count) template(s)  |  total time: $totalElapsed" -ForegroundColor Cyan
Write-Host $divider -ForegroundColor DarkCyan
Write-Host ''

$projWidth    = ($results | ForEach-Object { $_.ProjectName.Length } | Measure-Object -Maximum).Maximum
$projWidth    = [Math]::Max(20, [Math]::Min($projWidth, $Script:LineWidth - 50))
$statusWidth  = 20
$elapsedWidth = 10

$summaryHdr = '  {0}  {1}  {2}  {3}' -f 'Status'.PadRight($statusWidth), 'Elapsed'.PadRight($elapsedWidth), 'Project Name'.PadRight($projWidth), 'Execution ID'
$summarySep = '  {0}  {1}  {2}  {3}' -f ('-' * $statusWidth), ('-' * $elapsedWidth), ('-' * $projWidth), ('-' * 38)

Write-Host $summaryHdr -ForegroundColor White
Write-Host $summarySep -ForegroundColor DarkGray

$successCount = 0
$failCount    = 0

foreach ($r in $results) {
    if ($r.Status -in 'Succeeded', 'PartiallySucceeded') {
        $color = if ($r.Status -eq 'Succeeded') { 'Green' } else { 'Yellow' }
        $successCount++
    } else {
        $color = if ($r.Status -eq 'Skipped') { 'DarkGray' } else { 'Red' }
        $failCount++
    }

    $projTrunc = if ($r.ProjectName.Length -gt $projWidth) {
        $r.ProjectName.Substring(0, $projWidth - 3) + '...'
    } else { $r.ProjectName.PadRight($projWidth) }

    $row = '  {0}  {1}  {2}  {3}' -f $r.Status.PadRight($statusWidth), $r.Elapsed.PadRight($elapsedWidth), $projTrunc, $r.ExecutionId
    Write-Host $row -ForegroundColor $color
}

Write-Host ''
Write-Host $divider -ForegroundColor DarkCyan

$resultColor = if ($failCount -eq 0) { 'Green' } else { 'Red' }
$resultMsg   = if ($failCount -eq 0) {
    "  All $successCount template(s) exported successfully."
} else {
    "  $successCount succeeded  |  $failCount failed."
}
Write-Host $resultMsg -ForegroundColor $resultColor
Write-Host $divider   -ForegroundColor DarkCyan
Write-Host ''

Stop-RunTranscript

if ($failCount -gt 0) { exit 1 }
