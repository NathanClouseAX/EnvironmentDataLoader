#Requires -Version 5.1
<#
.SYNOPSIS
    Builds DMF export projects from selected templates and exports them.

.DESCRIPTION
    1. Authenticates once via Microsoft Entra device code flow.
    2. Retrieves the list of validated DMF definition-group templates.
    3. Presents an interactive selection menu (skipped when -TemplateName is supplied).
    4. Prompts for the source legal entity (skipped when -LegalEntityId is supplied).
    5. For each selected template:
       a. Fetches all template lines (DefinitionGroupTemplateLines).
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

    -WhatIf note:
      When -TemplateName is also supplied, no API calls are made at all.
      Without -TemplateName, the template list is fetched (read-only) so the
      interactive menu can be displayed; no projects are created or exported.

    All REST calls include automatic retry with exponential back-off (HTTP 5xx /
    429 / transient network errors).  The token expiry window is tracked and
    surfaced as a warning when approaching expiry.

    Library files (in ./lib/)
    ─────────────────────────
    DmfOutput.ps1   -- Write-* helpers, Format-Elapsed, Stop-RunTranscript
    DmfRequest.ps1  -- Invoke-DmfRequest (REST client with retry)

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

.PARAMETER DownloadPath
    Directory to save the downloaded export .zip files.
    Defaults to $env:TEMP.  Pass an empty string ('') to skip downloading.

.PARAMETER LogPath
    Path for the run transcript log.  When omitted, a log is auto-generated in
    DownloadPath (or $env:TEMP) as DMFProjectExport_<timestamp>.log.
    Pass an empty string ('') to suppress transcript logging entirely.

.PARAMETER PollIntervalSeconds
    How often (in seconds) to check export status.  Range: 5-300.  Default: 30.

.PARAMETER TimeoutMinutes
    Per-template polling timeout in minutes.  Range: 1-480.  Default: 60.

.PARAMETER MaxRetries
    Maximum number of automatic retries for transient REST failures.
    Range: 0-10.  Default: 3.

.PARAMETER Force
    Skip the confirmation prompt before exporting.

.PARAMETER WhatIf
    Show what would be created and exported without making any changes to D365.
    When -TemplateName is also provided, no API calls are made at all.
    Otherwise the template list is fetched (read-only) to populate the menu.

.PARAMETER PassThru
    Return per-template result objects to the pipeline after completion.
    Each object has properties: Template, TemplateId, ProjectName, LegalEntityId,
    Status, LinesAdded, ExecutionId, DownloadUrl, DownloadedTo, Elapsed.

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

    [AllowEmptyString()]
    [string]$LogPath,

    [ValidateRange(5, 300)]
    [int]$PollIntervalSeconds = 30,

    [ValidateRange(1, 480)]
    [int]$TimeoutMinutes = 60,

    [ValidateRange(0, 10)]
    [int]$MaxRetries = 3,

    [switch]$Force,
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
. (Join-Path $libPath 'DmfRequest.ps1')

Add-Type -AssemblyName System.IO.Compression.FileSystem

# =============================================================================
#  Pre-flight path validation  (before transcript so errors surface cleanly)
# =============================================================================
if ($DownloadPath -ne '' -and -not (Test-Path $DownloadPath -PathType Container)) {
    throw "Download path not found: '$DownloadPath'"
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
if ($DownloadPath -ne '') { Write-Info "Download to  : $DownloadPath" }
else                      { Write-Info 'Download     : disabled' }
if ($Script:TranscriptActive) { Write-Info "Log          : $LogPath" }
if ($WhatIf) { Write-Warn 'WhatIf active -- no projects will be created or exported.' }
if ($Force)  { Write-Info 'Force        : confirmation prompt suppressed' }

# =============================================================================
#  3.  WhatIf early exit for named template  (zero API calls)
# =============================================================================
if ($WhatIf -and $TemplateName) {
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
Write-Step 'Authenticating with Microsoft Entra (device code flow)'

$clientId    = '1950a258-227b-4e31-a9cf-717495945fc2'
$baseUrl     = $EnvironmentUrl.TrimEnd('/')
$authBase    = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0"
$scope       = "$baseUrl/.default"
$accessToken = $null
$tokenExpiry = [DateTime]::MaxValue

$deviceCode = Invoke-DmfRequest -Operation 'device code request' -Params @{
    Method      = 'Post'
    Uri         = "$authBase/devicecode"
    ContentType = 'application/x-www-form-urlencoded'
    Body        = "client_id=$clientId&scope=$([System.Uri]::EscapeDataString($scope))"
}

Write-Host ''
Write-Host $deviceCode.message -ForegroundColor Yellow
Write-Host ''

$pollUntil = (Get-Date).AddSeconds($deviceCode.expires_in)

while ((Get-Date) -lt $pollUntil) {
    Start-Sleep -Seconds $deviceCode.interval
    try {
        $tokenResp   = Invoke-RestMethod -Method Post `
            -Uri         "$authBase/token" `
            -ContentType 'application/x-www-form-urlencoded' `
            -Body        "grant_type=urn:ietf:params:oauth:grant-type:device_code&client_id=$clientId&device_code=$($deviceCode.device_code)"
        $accessToken = $tokenResp.access_token
        $tokenExpiry = (Get-Date).AddSeconds($tokenResp.expires_in - 60)
        Write-Info "Sign-in successful.  Token valid until $($tokenExpiry.ToString('HH:mm:ss'))."
        break
    }
    catch {
        $errCode = $null
        try { $errCode = ($_.ErrorDetails.Message | ConvertFrom-Json).error } catch {}
        switch ($errCode) {
            'authorization_pending'  { continue }
            'authorization_declined' { throw 'Sign-in declined.  Re-run and approve the prompt.' }
            'expired_token'          { throw 'Device code expired.  Re-run the script.' }
            default                  { throw }
        }
    }
}

if (-not $accessToken) { throw 'Authentication timed out before sign-in completed.' }

$authHeaders = @{ Authorization = "Bearer $accessToken" }
$dmfBase     = "$baseUrl/data/DataManagementDefinitionGroups/Microsoft.Dynamics.DataEntities"

# =============================================================================
#  5.  Fetch available templates  (OData with pagination)
# =============================================================================
Write-Step 'Fetching available templates'

$fetchUrl     = "$baseUrl/data/DefinitionGroupTemplateHeaders"
$rawTemplates = [System.Collections.Generic.List[psobject]]::new()

do {
    $page = Invoke-DmfRequest -Operation 'list templates' -Params @{
        Method  = 'Get'
        Uri     = $fetchUrl
        Headers = $authHeaders
    }
    foreach ($t in $page.value) { $rawTemplates.Add($t) }
    $fetchUrl = if ($page.PSObject.Properties['@odata.nextLink']) { $page.'@odata.nextLink' } else { $null }
} while ($fetchUrl)

# Only surface Validated templates; sort by TemplateId client-side.
$allTemplates = @($rawTemplates |
    Where-Object { $_.Status -eq 'Validated' } |
    Sort-Object  -Property TemplateId |
    ForEach-Object -Begin { $i = 1 } -Process {
    [pscustomobject]@{
        Index       = $i++
        TemplateId  = $_.TemplateId
        Description = if ($_.PSObject.Properties['Description'] -and $_.Description) { $_.Description } else { '' }
        ValidatedOn = if ($_.PSObject.Properties['ValidatedDateTime'] -and $_.ValidatedDateTime) {
                          try { ([datetime]$_.ValidatedDateTime).ToString('yyyy-MM-dd') } catch { '' }
                      } else { '' }
    }
})

if ($allTemplates.Count -eq 0) {
    throw 'No validated templates found in this environment.  Ensure DMF definition group templates exist and have Status = Validated.'
}

Write-Info "$($allTemplates.Count) validated template(s) found."

# =============================================================================
#  6.  Template selection
# =============================================================================
$selectedTemplates = [System.Collections.Generic.List[pscustomobject]]::new()

if ($TemplateName) {
    # ── Single-template mode (non-interactive) ──────────────────────────────
    $match = $allTemplates | Where-Object { $_.TemplateId -eq $TemplateName }
    if (-not $match) {
        $available = ($allTemplates | ForEach-Object { "    '$($_.TemplateId)'" }) -join [System.Environment]::NewLine
        throw "Template '$TemplateName' not found in environment '$EnvironmentUrl'.`nAvailable templates:`n$available"
    }
    $selectedTemplates.Add($match)
    Write-Info "Template : $TemplateName"
} else {
    # ── Interactive numbered selection ─────────────────────────────────────
    Write-Step 'Select templates to export'

    $idWidth   = ($allTemplates | ForEach-Object { $_.TemplateId.Length }   | Measure-Object -Maximum).Maximum
    $idWidth   = [int][Math]::Max(20, [Math]::Min($idWidth, [int]([Math]::Floor(($Script:LineWidth - 22) * 0.45))))
    $descWidth = [int][Math]::Max(20, [Math]::Min(
        ($allTemplates | ForEach-Object { $_.Description.Length } | Measure-Object -Maximum).Maximum,
        $Script:LineWidth - $idWidth - 22
    ))

    $colFmt = '  {0,3}  {1}  {2}  {3}'
    Write-Host ''
    Write-Host ($colFmt -f '#', 'Template ID'.PadRight($idWidth), 'Description'.PadRight($descWidth), 'Validated') -ForegroundColor White
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
        if ($tmpl.ValidatedOn)  { $meta += "  [validated: $($tmpl.ValidatedOn)]" }
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
Write-Rule 'Ready to export'
Write-Host ''
Write-Info "Environment  : $EnvironmentUrl"
Write-Info "Legal entity : $LegalEntityId"
Write-Info "Templates    : $($selectedTemplates.Count)"

foreach ($tmpl in $selectedTemplates) {
    $meta = "$($tmpl.TemplateId) $LegalEntityId"
    if ($tmpl.Description) { $meta += "  -- $($tmpl.Description)" }
    if ($tmpl.ValidatedOn) { $meta += "  [validated: $($tmpl.ValidatedOn)]" }
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

# =============================================================================
#  10.  Process templates
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

    # Token expiry warnings
    if ((Get-Date) -ge $tokenExpiry) {
        Write-Warn 'Access token has expired.  API calls will likely fail with HTTP 401.  Re-run the script.'
    } elseif ((Get-Date).AddMinutes(5) -ge $tokenExpiry) {
        Write-Warn "Token expires at $($tokenExpiry.ToString('HH:mm:ss')) -- it may expire mid-export."
    }

    try {
        # ── a. Fetch template lines ───────────────────────────────────────────
        Write-Step "Fetching template lines  ($tmplId)"

        # OData single-quote escape (double any single quotes in the value)
        $oDataSafeId    = $tmplId.Replace("'", "''")
        $encodedFilter  = [System.Uri]::EscapeDataString("TemplateId eq '$oDataSafeId'")
        $linesUrl       = "$baseUrl/data/DefinitionGroupTemplateLines?`$filter=$encodedFilter"
        $rawLines       = [System.Collections.Generic.List[psobject]]::new()

        do {
            $linesPage = Invoke-DmfRequest -Operation 'list template lines' -Params @{
                Method  = 'Get'
                Uri     = $linesUrl
                Headers = $authHeaders
            }
            foreach ($ln in $linesPage.value) { $rawLines.Add($ln) }
            $linesUrl = if ($linesPage.PSObject.Properties['@odata.nextLink']) { $linesPage.'@odata.nextLink' } else { $null }
        } while ($linesUrl)

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
            Invoke-DmfRequest -Operation 'DeleteProject' -Params @{
                Method  = 'Delete'
                Uri     = "$baseUrl/data/DataManagementDefinitionGroups('$encodedName')"
                Headers = $authHeaders
            }
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
                        # Download zip.  Suppress progress bar -- Invoke-WebRequest renders
                        # progress on every packet by default, making large downloads very slow in PS5.1.
                        $prevPref = $ProgressPreference
                        $ProgressPreference = 'SilentlyContinue'
                        try {
                            Invoke-WebRequest -Uri $downloadUrl -OutFile $zipFile -UseBasicParsing
                        } finally {
                            $ProgressPreference = $prevPref
                        }
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
