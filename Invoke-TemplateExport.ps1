#Requires -Version 5.1
<#
.SYNOPSIS
    Lists DMF templates from a D365 F&O environment and creates export projects
    for selected templates in the specified legal entity.

.DESCRIPTION
    1. Authenticates once via Microsoft Entra device code flow.
    2. Retrieves the list of available DMF definition-group templates from D365.
    3. Presents an interactive selection menu (skipped when -TemplateName is supplied).
    4. Prompts for confirmation before exporting (suppressed with -Force).
    5. For each confirmed template:
       a. Submits an ExportToPackage job to D365 DMF.
       b. Polls until the job reaches a terminal status, showing a live progress bar.
       c. Retrieves the blob download URL for the exported package.
       d. Downloads the package to -DownloadPath when that parameter is non-empty.
    6. Prints a colour-coded summary table with per-template elapsed time and
       execution IDs for follow-up in D365 Job history.
    7. Writes a timestamped transcript to -LogPath (auto-generated when omitted;
       pass an empty string to suppress transcript logging entirely).

    -WhatIf note:
      When -TemplateName is also supplied, no API calls are made at all.
      Without -TemplateName, the template list is fetched (read-only) so the
      interactive menu can be displayed; no export jobs are submitted.

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

.PARAMETER EnvironmentUrl
    Base URL of the D365 F&O environment (no trailing slash).
    Example: https://contoso.operations.dynamics.com

.PARAMETER TenantId
    Microsoft Entra tenant ID (GUID) or verified domain name.
    Example: contoso.onmicrosoft.com

.PARAMETER LegalEntityId
    D365 legal entity (company) to export from.
    Example: DAT

.PARAMETER TemplateName
    When supplied, exports this one template without prompting.
    The value must match a TemplateId exactly.
    Example: 'MyExportTemplate'

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
    DownloadPath (or $env:TEMP) as DMFExport_<timestamp>.log.
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
    Show what would be exported without submitting any jobs to D365.
    When -TemplateName is also provided, no API calls are made at all.
    Otherwise the template list is fetched (read-only) to populate the menu.

.PARAMETER PassThru
    Return per-template result objects to the pipeline after completion.
    Each object has properties: Template, TemplateId, Status, ExecutionId,
    DownloadUrl, DownloadedTo, Elapsed.

.EXAMPLE
    # Interactive: list all templates and prompt for selection
    .\Invoke-TemplateExport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'DAT'

.EXAMPLE
    # Non-interactive: export one template, skip confirm, download to C:\Exports
    .\Invoke-TemplateExport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'DAT' `
        -TemplateName   'SystemSetupExport' `
        -DownloadPath   'C:\Exports' `
        -Force

.EXAMPLE
    # WhatIf with a named template -- no API calls at all
    .\Invoke-TemplateExport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'DAT' `
        -TemplateName   'SystemSetupExport' `
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

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$LegalEntityId,

    [string]$TemplateName,

    [AllowEmptyString()]
    [string]$DownloadPath = $env:TEMP,

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

# =============================================================================
#  Pre-flight path validation  (before transcript so errors surface cleanly)
# =============================================================================
# The download folder is output: create it rather than demanding it exists.
if ($DownloadPath -ne '' -and -not (Test-Path $DownloadPath -PathType Container)) {
    if ($WhatIf) { Write-Host "Download path '$DownloadPath' does not exist; it would be created." -ForegroundColor Yellow }
    else         { New-Item -ItemType Directory -Path $DownloadPath -Force | Out-Null }
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
    $LogPath = Join-Path $logBase "DMFExport_${ts}.log"
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
Write-Banner -DryRun:$WhatIf -Title 'D365 F&O Template Export Utility'
Write-Host ''
Write-Info "Environment  : $EnvironmentUrl"
Write-Info "Tenant       : $TenantId"
Write-Info "Legal entity : $LegalEntityId"
if ($DownloadPath -ne '') { Write-Info "Download to  : $DownloadPath" }
else                      { Write-Info 'Download     : disabled' }
if ($Script:TranscriptActive) { Write-Info "Log          : $LogPath" }
if ($WhatIf) { Write-Warn 'WhatIf active -- no export jobs will be submitted.' }
if ($Force)  { Write-Info 'Force        : confirmation prompt suppressed' }

# =============================================================================
#  3.  WhatIf early exit for named template  (zero API calls)
# =============================================================================
if ($WhatIf -and $TemplateName) {
    Write-Host ''
    Write-Rule 'WhatIf -- no export will occur'
    Write-Host ''
    Write-Info "Would export template : $TemplateName"
    Write-Info "Target legal entity   : $LegalEntityId"
    Write-Host ''
    Write-Warn 'Template name is not validated in WhatIf mode; re-run without -WhatIf to confirm it exists.'
    Write-Info 'No API calls made.  Remove -WhatIf to perform the export.'
    if ($PassThru) {
        [pscustomobject]@{
            Template     = $TemplateName
            TemplateId   = $TemplateName
            Status       = 'WhatIf'
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

$baseUrl = $EnvironmentUrl.TrimEnd('/')
$session = Connect-DmfEnvironment -EnvironmentUrl $baseUrl -TenantId $TenantId -AuthMode $AuthMode

# Every Invoke-DmfRequest call to this environment reads the session and
# renews the token silently when it is close to expiry (lib/DmfAuth.ps1).
$Script:DmfSession = $session
$authHeaders       = Get-DmfAuthHeaders -Session $session
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
    throw 'No templates found in this environment.  Ensure DMF definition group templates exist (DefinitionGroupTemplateHeaders).'
}

Write-Info "$($allTemplates.Count) template(s) found."

# =============================================================================
#  6.  Template selection
# =============================================================================
$selectedTemplates = [System.Collections.Generic.List[pscustomobject]]::new()

if ($TemplateName) {
    # ── Single-template mode (non-interactive) ──────────────────────────────
    $match = $allTemplates | Where-Object { $_.TemplateId -eq $TemplateName }
    if (-not $match) {
        # ExportToPackage resolves definitionGroupId against data projects, not
        # templates, so a project name (for example the "<template> <company>"
        # project that Invoke-ProjectExport.ps1 creates) is just as valid here.
        $encodedName = [System.Uri]::EscapeDataString($TemplateName)
        $project = $null
        try {
            $project = Invoke-DmfRequest -Operation 'CheckProject' -Params @{
                Method  = 'Get'
                Uri     = "$baseUrl/data/DataManagementDefinitionGroups('$encodedName')"
                Headers = $authHeaders
            }
        } catch {
            if ($_ -notmatch 'HTTP 404' -and $_ -notmatch 'HTTP 400') { throw }
        }
        if ($null -ne $project) {
            $desc  = if ($project.PSObject.Properties['Description'] -and $project.Description) { [string]$project.Description } else { '' }
            $match = [pscustomobject]@{ Index = 0; TemplateId = $TemplateName; Description = $desc; ValidatedOn = 'data project' }
            Write-Info "'$TemplateName' is a data project (not a template); exporting it directly."
        }
    }
    if (-not $match) {
        $available = ($allTemplates | ForEach-Object { "    '$($_.TemplateId)'" }) -join [System.Environment]::NewLine
        throw "'$TemplateName' is neither a validated template nor a data project in '$EnvironmentUrl'.`nAvailable templates:`n$available"
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
#  7.  WhatIf exit  (interactive path -- list was already fetched read-only)
# =============================================================================
if ($WhatIf) {
    Write-Host ''
    Write-Rule 'WhatIf -- no export will occur'
    Write-Host ''
    Write-Info "Environment  : $EnvironmentUrl"
    Write-Info "Legal entity : $LegalEntityId"
    Write-Info "Templates    : $($selectedTemplates.Count)"
    foreach ($tmpl in $selectedTemplates) {
        $meta = $tmpl.TemplateId
        if ($tmpl.Description)  { $meta += "  -- $($tmpl.Description)" }
        if ($tmpl.ValidatedOn)  { $meta += "  [validated: $($tmpl.ValidatedOn)]" }
        Write-Detail "[$($tmpl.Index)] $meta"
    }
    Write-Host ''
    Write-Info 'No export jobs submitted.  Remove -WhatIf to perform the export.'
    if ($PassThru) {
        $selectedTemplates | ForEach-Object {
            [pscustomobject]@{
                Template     = $_.Description
                TemplateId   = $_.TemplateId
                Status       = 'WhatIf'
                ExecutionId  = '-'
                DownloadUrl  = '-'
                DownloadedTo = '-'
                Elapsed      = '-'
            }
        }
    }
    Stop-RunTranscript
    exit 0
}

# =============================================================================
#  8.  Confirmation  (skipped with -Force)
# =============================================================================
Write-Host ''
Write-Rule 'Ready to export'
Write-Host ''
Write-Info "Environment  : $EnvironmentUrl"
Write-Info "Legal entity : $LegalEntityId"
Write-Info "Templates    : $($selectedTemplates.Count)"

foreach ($tmpl in $selectedTemplates) {
    $meta = $tmpl.TemplateId
    if ($tmpl.Description)  { $meta += "  -- $($tmpl.Description)" }
    if ($tmpl.ValidatedOn)  { $meta += "  [validated: $($tmpl.ValidatedOn)]" }
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
#  9.  Process templates
# =============================================================================
$results     = [System.Collections.Generic.List[pscustomobject]]::new()
$scriptStart = Get-Date

foreach ($tmpl in $selectedTemplates) {
    $tmplId      = $tmpl.TemplateId
    $tmplSafe    = ($tmplId -replace '[^A-Za-z0-9]', '-') -replace '-{2,}', '-'
    $tmplStart   = Get-Date
    $timestamp   = Get-Date -Format 'yyyyMMddHHmmss'
    $tmplResult  = [pscustomobject]@{
        Template     = $tmpl.Description
        TemplateId   = $tmplId
        Status       = 'Error'
        ExecutionId  = '-'
        DownloadUrl  = '-'
        DownloadedTo = '-'
        Elapsed      = '-'
    }

    $divider = '=' * $Script:LineWidth
    Write-Host ''
    Write-Host $divider -ForegroundColor DarkCyan
    Write-Host "  [$($results.Count + 1)/$($selectedTemplates.Count)]  $tmplId" -ForegroundColor Cyan
    if ($tmpl.Description) {
        Write-Host "  $($tmpl.Description)" -ForegroundColor DarkGray
    }
    Write-Host $divider -ForegroundColor DarkCyan

    # Token expiry: renewed silently when a refresh token is available,
    # otherwise warned about as before.
    [void](Test-DmfTokenExpiry -Session $session -Activity 'export')

    try {
        # ── a. Submit export job ──────────────────────────────────────────────
        $packageName = "${tmplSafe}_${timestamp}"
        Write-Step "Submitting export job  (company: $LegalEntityId)"
        Write-Info "Package name : $packageName"

        $exportResp = Invoke-DmfRequest -Operation 'ExportToPackage' -Params @{
            Method      = 'Post'
            Uri         = "$dmfBase.ExportToPackage"
            Headers     = $authHeaders
            ContentType = 'application/json'
            Body        = (@{
                definitionGroupId = $tmplId
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

        # ── b. Poll for completion ────────────────────────────────────────────
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

        # ── c. Retrieve download URL  (on success only) ───────────────────────
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

                # ── d. Download package ───────────────────────────────────────
                if ($DownloadPath -ne '') {
                    Write-Step 'Downloading exported package'
                    $fileName   = "${tmplSafe}_${timestamp}.zip"
                    $outputFile = Join-Path $DownloadPath $fileName
                    try {
                        # Routed through Invoke-DmfDownload so a throttled or
                        # flaky blob endpoint backs off and resumes instead of
                        # failing the whole export.
                        Invoke-DmfDownload -Uri $downloadUrl -OutFile $outputFile -Operation 'download package'
                        $fileSizeMB = [Math]::Round((Get-Item $outputFile).Length / 1MB, 2)
                        Write-OK "Downloaded : $outputFile  ($fileSizeMB MB)"
                        $tmplResult.DownloadedTo = $outputFile
                    } catch {
                        Write-Warn "Download failed: $_"
                        Write-Info 'Use the download URL above to retrieve the package manually.'
                    }
                }
            } else {
                Write-Warn 'GetExportedPackageUrl returned an empty URL.'
            }
        }

        # ── e. Package outcome message ────────────────────────────────────────
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
#  10.  Summary
# =============================================================================
$totalElapsed = Format-Elapsed ((Get-Date) - $scriptStart)
$divider      = '=' * $Script:LineWidth

Write-Host ''
Write-Host $divider -ForegroundColor DarkCyan
Write-Host "  Summary  --  $($results.Count) template(s)  |  total time: $totalElapsed" -ForegroundColor Cyan
Write-Host $divider -ForegroundColor DarkCyan
Write-Host ''

$idWidth2     = ($results | ForEach-Object { $_.TemplateId.Length } | Measure-Object -Maximum).Maximum
$idWidth2     = [Math]::Max(20, [Math]::Min($idWidth2, $Script:LineWidth - 55))
$statusWidth  = 20
$elapsedWidth = 10

$summaryHdr = '  {0}  {1}  {2}  {3}' -f 'Status'.PadRight($statusWidth), 'Elapsed'.PadRight($elapsedWidth), 'Template ID'.PadRight($idWidth2), 'Execution ID'
$summarySep = '  {0}  {1}  {2}  {3}' -f ('-' * $statusWidth), ('-' * $elapsedWidth), ('-' * $idWidth2), ('-' * 38)

Write-Host $summaryHdr -ForegroundColor White
Write-Host $summarySep -ForegroundColor DarkGray

$successCount = 0
$failCount    = 0

foreach ($r in $results) {
    if ($r.Status -in 'Succeeded', 'PartiallySucceeded') {
        $color = if ($r.Status -eq 'Succeeded') { 'Green' } else { 'Yellow' }
        $successCount++
    } else {
        $color = 'Red'
        $failCount++
    }

    $idTrunc = if ($r.TemplateId.Length -gt $idWidth2) {
        $r.TemplateId.Substring(0, $idWidth2 - 3) + '...'
    } else { $r.TemplateId.PadRight($idWidth2) }

    $row = '  {0}  {1}  {2}  {3}' -f $r.Status.PadRight($statusWidth), $r.Elapsed.PadRight($elapsedWidth), $idTrunc, $r.ExecutionId
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
