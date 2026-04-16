#Requires -Version 5.1
<#
.SYNOPSIS
    Reports on D365 F&O Data Management "Copy legal entity" execution jobs and
    produces a professional HTML report of per-entity staging and target results.

.DESCRIPTION
    1. Authenticates once via Microsoft Entra device code flow.
    2. Fetches all DataManagementExecutionJobs from the D365 OData endpoint,
       with automatic page-following for large result sets.
    3. Filters to jobs whose Description begins with "Copy legal entity"
       (e.g. "Copy legal entity data for abmc to ABUS"), case-insensitive.
    4. For each matching job, fetches all DataManagementExecutionJobDetails records.
    5. For jobs that contain any entity-level error or aborted status, calls the
       D365 Data Management API (GetExecutionSummaryStatus) to obtain the
       platform-level execution result as an additional diagnostic detail.
    6. Renders a colour-coded console table and a self-contained HTML report.
         Green  -- Finished with no failed records (no action required)
         Yellow -- Partial issue: non-error anomaly or failed-record count > 0
         Red    -- Staging or target error / aborted (manual intervention required)
    7. Optionally restricts the displayed rows to records with issues (-IssuesOnly).
    8. Prints a run summary to the console and writes the HTML report to -HtmlPath.
    9. Writes a timestamped transcript to -LogPath (auto-generated when omitted;
       pass an empty string to suppress transcript logging entirely).

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

.PARAMETER LogPath
    Path for the run transcript log.  When omitted, a log is auto-generated in
    $env:TEMP as DMFJobReport_<timestamp>.log.  Pass an empty string ('') to
    suppress transcript logging entirely.

.PARAMETER HtmlPath
    Path for the HTML report file.  When omitted, a report is auto-generated in
    $env:TEMP as DMFJobReport_<timestamp>.html.  Pass an empty string ('') to
    suppress HTML output entirely.

.PARAMETER MaxRetries
    Maximum number of automatic retries for transient REST failures.
    Range: 0-10.  Default: 3.

.PARAMETER IssuesOnly
    When set, restricts both the console table and the HTML report to entity
    records that need attention: any staging or target status that is not
    Finished/NotRun, or a failed-record count greater than zero.  Clean entities
    are still counted in the summary cards.

.PARAMETER PassThru
    Return the per-entity enriched detail objects to the pipeline after the
    report is displayed.  Each object has properties: JobId, JobDescription,
    DefinitionGroupId, EntityName, StagingStatus, TargetStatus,
    StagingRecordsToBeProcessedCount, TargetRecordsCreatedCount,
    TargetRecordsUpdatedCount, FailedRecords, NeedsAttention.

.EXAMPLE
    # Full report -- all entities for all matching jobs
    .\Get-ExecutionJobReport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com'

.EXAMPLE
    # Show only records that require attention; write HTML to a specific path
    .\Get-ExecutionJobReport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -IssuesOnly `
        -HtmlPath       'C:\Reports\job_report.html'

.EXAMPLE
    # Pipe results to a CSV for external tracking
    .\Get-ExecutionJobReport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -PassThru |
        Export-Csv -Path 'C:\Reports\job_report.csv' -NoTypeInformation
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^https?://')]
    [string]$EnvironmentUrl,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [AllowEmptyString()]
    [string]$LogPath,

    [AllowEmptyString()]
    [string]$HtmlPath,

    [ValidateRange(0, 10)]
    [int]$MaxRetries = 3,

    [switch]$IssuesOnly,

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

# =============================================================================
#  Script-level constants  (consumed by lib functions via $Script: scope)
# =============================================================================
$Script:Version          = '2.0'
$Script:LineWidth        = 80
$Script:MaxRetries       = $MaxRetries
$Script:TranscriptActive = $false

try {
    $w = $Host.UI.RawUI.BufferSize.Width
    $Script:LineWidth = [Math]::Min(140, [Math]::Max(80, $w))
} catch { <# non-interactive host -- keep default #> }

# =============================================================================
#  HTML helper functions
# =============================================================================
function ConvertTo-HtmlEncoded {
    param([string]$Text)
    if (-not $Text) { return '' }
    $Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;'
}

function Get-StatusBadgeHtml {
    param([string]$Status)
    $label = if ($Status) { ConvertTo-HtmlEncoded $Status } else { 'N/A' }
    $class = switch ($Status) {
        'Finished'           { 'b-ok'   }
        'Error'              { 'b-err'  }
        'Aborted'            { 'b-err'  }
        'PartiallySucceeded' { 'b-warn' }
        default              { 'b-none' }
    }
    "<span class='badge $class'>$label</span>"
}

function Get-LegalEntities {
    # Parses "Copy legal entity [data for] {source} to {target}" from a job description.
    param([string]$Description)
    if ($Description -match '(?i)copy legal entity(?:\s+data\s+for)?\s+(\w+)\s+to\s+(\w+)') {
        return [pscustomobject]@{ Source = $Matches[1].ToUpper(); Target = $Matches[2].ToUpper(); Found = $true }
    }
    return [pscustomobject]@{ Source = ''; Target = ''; Found = $false }
}

# =============================================================================
#  OData pagination helper
# =============================================================================
function Get-ODataAll {
    <#
    .SYNOPSIS  Fetches every page of an OData endpoint and emits each item.
    .NOTES     Caller collects via @(Get-ODataAll ...) to get a typed array.
    #>
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$Operation,
        [Parameter(Mandatory)][hashtable]$Headers
    )
    $page    = 0
    $nextUri = $Uri
    do {
        $page++
        $resp = Invoke-DmfRequest -Operation "$Operation (page $page)" -Params @{
            Method  = 'Get'
            Uri     = $nextUri
            Headers = $Headers
        }
        if ($resp.value) { $resp.value | Write-Output }
        $nextUri = if ($resp.PSObject.Properties['@odata.nextLink']) { $resp.'@odata.nextLink' } else { $null }
    } while ($nextUri)
}

# =============================================================================
#  1.  Transcript startup  (timestamp shared by log and HTML default paths)
# =============================================================================
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'

if (-not $PSBoundParameters.ContainsKey('LogPath')) {
    $LogPath = Join-Path $env:TEMP "DMFJobReport_${ts}.log"
}
if (-not $PSBoundParameters.ContainsKey('HtmlPath')) {
    $HtmlPath = Join-Path $env:TEMP "DMFJobReport_${ts}.html"
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
Write-Banner -Title 'D365 F&O DMF Execution Job Report'
Write-Host ''
Write-Info "Environment  : $EnvironmentUrl"
Write-Info "Tenant       : $TenantId"
if ($Script:TranscriptActive)  { Write-Info "Log          : $LogPath"  }
if ($HtmlPath -ne '')           { Write-Info "HTML report  : $HtmlPath" }
if ($IssuesOnly) { Write-Info 'IssuesOnly   : only records needing attention will appear in the report' }

# =============================================================================
#  3.  Authenticate  (once -- same device code flow used across all scripts)
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
        $tokenExpiry = (Get-Date).AddSeconds($tokenResp.expires_in - 60)   # 60 s safety buffer
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
$odataBase   = "$baseUrl/data"

# =============================================================================
#  4.  Fetch all execution jobs; filter to "Copy legal entity" descriptions
# =============================================================================
Write-Step 'Fetching DataManagementExecutionJobs'

$jobsUri     = "$odataBase/DataManagementExecutionJobs?`$select=JobId,Description"
$allJobs     = @(Get-ODataAll -Uri $jobsUri -Operation 'GetExecutionJobs' -Headers $authHeaders)

Write-Info "Total execution jobs returned : $($allJobs.Count)"

# Case-insensitive prefix match -- captures any "Copy legal entity ..." description.
$matchedJobs = @($allJobs | Where-Object { $_.Description -like 'Copy legal entity*' })

if ($matchedJobs.Count -eq 0) {
    Write-Warn "No execution jobs found whose Description begins with 'Copy legal entity'."
    Stop-RunTranscript
    exit 0
}

Write-Info "Jobs matching 'Copy legal entity*' : $($matchedJobs.Count)"
foreach ($j in ($matchedJobs | Sort-Object Description)) {
    Write-Detail "  $($j.Description)  [$($j.JobId)]"
}

# =============================================================================
#  5.  Fetch entity details for each matching job; call DMF API for errors
# =============================================================================
Write-Step 'Fetching DataManagementExecutionJobDetails'

$cleanStatuses = @('Finished', 'NotRun', 'None', '')
$errorStatuses = @('Error', 'Aborted')

$allDetails   = [System.Collections.Generic.List[pscustomobject]]::new()
$jobSummaries = [System.Collections.Generic.List[pscustomobject]]::new()

foreach ($job in ($matchedJobs | Sort-Object Description)) {
    if ((Get-Date) -ge $tokenExpiry) {
        Write-Warn 'Access token has expired.  Re-run the script to re-authenticate.'
        break
    }

    Write-Info "  Fetching details : $($job.Description)"

    $jobIdLiteral  = $job.JobId -replace "'", "''"
    $encodedFilter = [System.Uri]::EscapeDataString("JobId eq '$jobIdLiteral'")
    $detailsUri    = "$odataBase/DataManagementExecutionJobDetails?`$filter=$encodedFilter"

    $rawDetails = @(Get-ODataAll -Uri $detailsUri -Operation "GetJobDetails:$($job.Description)" -Headers $authHeaders)

    if ($rawDetails.Count -eq 0) {
        Write-Detail '    No entity detail records found.'
        continue
    }

    Write-Detail "    $($rawDetails.Count) entity record(s) found."

    $jobHasError = $false

    foreach ($d in $rawDetails) {
        $staged  = [int]($d.StagingRecordsToBeProcessedCount)
        $created = [int]($d.TargetRecordsCreatedCount)
        $updated = [int]($d.TargetRecordsUpdatedCount)
        $failed  = [Math]::Max(0, $staged - $created - $updated)

        $isError        = ($d.TargetStatus  -in $errorStatuses) -or
                          ($d.StagingStatus -in $errorStatuses)
        $needsAttention = $isError -or
                          ($d.TargetStatus  -notin $cleanStatuses) -or
                          ($d.StagingStatus -notin $cleanStatuses) -or
                          ($failed -gt 0)

        if ($isError) { $jobHasError = $true }

        $allDetails.Add([pscustomobject]@{
            JobId                            = $job.JobId
            JobDescription                   = $job.Description
            DefinitionGroupId                = $d.DefinitionGroupId
            EntityName                       = $d.EntityName
            StagingStatus                    = $d.StagingStatus
            TargetStatus                     = $d.TargetStatus
            StagingRecordsToBeProcessedCount = $staged
            TargetRecordsCreatedCount        = $created
            TargetRecordsUpdatedCount        = $updated
            FailedRecords                    = $failed
            NeedsAttention                   = $needsAttention
            IsError                          = $isError
        })
    }

    # Call GetExecutionSummaryStatus for jobs with entity-level errors to get
    # the platform-level D365 execution result as an additional data point.
    if ($jobHasError) {
        try {
            $summaryResp = Invoke-DmfRequest `
                -Operation "GetExecutionSummaryStatus:$($job.Description)" `
                -Params @{
                    Method      = 'Post'
                    Uri         = "$dmfBase.GetExecutionSummaryStatus"
                    Headers     = $authHeaders
                    ContentType = 'application/json'
                    Body        = (@{ executionId = $job.JobId } | ConvertTo-Json)
                }
            $apiStatus = $summaryResp.value
            Write-Detail "    API execution status : $apiStatus"
            $jobSummaries.Add([pscustomobject]@{
                JobId     = $job.JobId
                ApiStatus = $apiStatus
            })
        }
        catch {
            Write-Detail "    Could not retrieve API execution status: $_"
        }
    }
}

if ($allDetails.Count -eq 0) {
    Write-Warn 'No entity detail records found for any matching job.'
    Stop-RunTranscript
    exit 0
}

# =============================================================================
#  6.  Compute summary counts
# =============================================================================
$errorCount  = @($allDetails | Where-Object { $_.IsError }).Count
$warnCount   = @($allDetails | Where-Object { $_.NeedsAttention -and -not $_.IsError }).Count
$okCount     = $allDetails.Count - $errorCount - $warnCount
$totalFailed = [int]($allDetails | Measure-Object -Property FailedRecords -Sum).Sum

# =============================================================================
#  7.  Console table
# =============================================================================
$displaySet = @(if ($IssuesOnly) {
    $allDetails | Where-Object { $_.NeedsAttention } | Sort-Object DefinitionGroupId, EntityName
} else {
    $allDetails | Sort-Object DefinitionGroupId, EntityName
})

$divider = '=' * $Script:LineWidth

Write-Host ''
Write-Host $divider -ForegroundColor DarkCyan
$filterTag = if ($IssuesOnly) { '  [Issues Only]' } else { '' }
Write-Host "  DMF Execution Job Detail Report$filterTag" -ForegroundColor Cyan
Write-Host "  Jobs: $($matchedJobs.Count)  |  Total entities: $($allDetails.Count)  |  Displayed: $($displaySet.Count)" -ForegroundColor Cyan
Write-Host $divider -ForegroundColor DarkCyan
Write-Host ''

if ($displaySet.Count -eq 0) {
    Write-OK 'No records to display.  All entity records completed without issues.'
} else {
    $grpW = [Math]::Min(28, [Math]::Max(15, ($displaySet | ForEach-Object { $_.DefinitionGroupId.Length } | Measure-Object -Maximum).Maximum))
    $entW = [Math]::Min(38, [Math]::Max(15, ($displaySet | ForEach-Object { $_.EntityName.Length }        | Measure-Object -Maximum).Maximum))
    $stW  = 10
    $tgW  = 10
    $nmW  = 7

    $colFmt = '  {0}  {1}  {2}  {3}  {4}  {5}  {6}  {7}'
    Write-Host ($colFmt -f 'Group'.PadRight($grpW), 'Entity'.PadRight($entW), 'Staging'.PadRight($stW), 'Target'.PadRight($tgW), 'Staged'.PadLeft($nmW), 'Created'.PadLeft($nmW), 'Updated'.PadLeft($nmW), 'Failed'.PadLeft($nmW)) -ForegroundColor White
    Write-Host ($colFmt -f ('-'*$grpW), ('-'*$entW), ('-'*$stW), ('-'*$tgW), ('-'*$nmW), ('-'*$nmW), ('-'*$nmW), ('-'*$nmW)) -ForegroundColor DarkGray

    $lastGroup = $null
    foreach ($d in $displaySet) {
        if ($null -ne $lastGroup -and $d.DefinitionGroupId -ne $lastGroup) { Write-Host '' }
        $lastGroup = $d.DefinitionGroupId

        $grp   = if ($d.DefinitionGroupId.Length -gt $grpW) { $d.DefinitionGroupId.Substring(0, $grpW-3)+'...' } else { $d.DefinitionGroupId.PadRight($grpW) }
        $ent   = if ($d.EntityName.Length         -gt $entW) { $d.EntityName.Substring(0, $entW-3)+'...'        } else { $d.EntityName.PadRight($entW)         }
        $color = if ($d.IsError) { 'Red' } elseif ($d.NeedsAttention) { 'Yellow' } else { 'Green' }

        Write-Host ($colFmt -f $grp, $ent, $d.StagingStatus.PadRight($stW), $d.TargetStatus.PadRight($tgW), $d.StagingRecordsToBeProcessedCount.ToString().PadLeft($nmW), $d.TargetRecordsCreatedCount.ToString().PadLeft($nmW), $d.TargetRecordsUpdatedCount.ToString().PadLeft($nmW), $d.FailedRecords.ToString().PadLeft($nmW)) -ForegroundColor $color
    }
}

# =============================================================================
#  8.  Console summary
# =============================================================================
Write-Host ''
Write-Host $divider -ForegroundColor DarkCyan
Write-Host '  Summary' -ForegroundColor Cyan
Write-Host $divider -ForegroundColor DarkCyan
Write-Host ''
Write-Host "  Jobs matched     : $($matchedJobs.Count)" -ForegroundColor Gray
Write-Host "  Total entities   : $($allDetails.Count)"  -ForegroundColor Gray
Write-Host ''
Write-Host "    Clean (no action)  : $okCount"     -ForegroundColor $(if ($okCount -eq $allDetails.Count) { 'Green'  } else { 'Gray'   })
Write-Host "    Warnings           : $warnCount"   -ForegroundColor $(if ($warnCount  -gt 0)              { 'Yellow' } else { 'Gray'   })
Write-Host "    Errors             : $errorCount"  -ForegroundColor $(if ($errorCount -gt 0)              { 'Red'    } else { 'Gray'   })
Write-Host "    Total failed recs  : $totalFailed" -ForegroundColor $(if ($totalFailed -gt 0)             { 'Yellow' } else { 'Gray'   })

if ($errorCount -gt 0 -or $warnCount -gt 0) {
    Write-Host ''
    Write-Warn 'One or more entities require manual intervention.'
    Write-Info 'Correct source data for the highlighted entities, then re-run the import.'
    if (-not $IssuesOnly) { Write-Info 'Tip: re-run with -IssuesOnly to display only records that need attention.' }
}

Write-Host ''
Write-Host $divider -ForegroundColor DarkCyan
Write-Host ''

# =============================================================================
#  9.  HTML report
# =============================================================================
if ($HtmlPath -ne '') {
    Write-Step 'Generating HTML report'

    $reportTimestamp = Get-Date -Format 'MMMM d, yyyy  h:mm tt'
    $filterNoticeHtml = if ($IssuesOnly) {
        '<div class="filter-notice">&#9432;&nbsp; This report shows only entities that require attention.  Clean records are included in the summary counts above but are not shown in the table.</div>'
    } else { '' }

    # -- Build table body: one job-group header row per job, then entity rows --
    $tbodySb = [System.Text.StringBuilder]::new()

    $jobGroups = @($allDetails | Group-Object JobId | Sort-Object { $_.Group[0].JobDescription })

    foreach ($jg in $jobGroups) {
        # Force to a typed PS array so .Count is always available under strict mode.
        $groupItems = @($jg.Group)

        $jobDesc   = $groupItems[0].JobDescription
        $companies = Get-LegalEntities -Description $jobDesc

        $jgErrors  = @($groupItems | Where-Object { $_.IsError }).Count
        $jgWarns   = @($groupItems | Where-Object { $_.NeedsAttention -and -not $_.IsError }).Count
        $jgFailed  = [int]($groupItems | Measure-Object FailedRecords -Sum).Sum

        # API summary status for this job (if fetched)
        $summaryObj = $jobSummaries | Where-Object { $_.JobId -eq $jg.Name } | Select-Object -First 1
        $apiStatus  = if ($summaryObj) { $summaryObj.ApiStatus } else { '' }

        $jobRowClass = if ($jgErrors -gt 0) { 'job-row has-errors' } elseif ($jgWarns -gt 0) { 'job-row has-warnings' } else { 'job-row' }
        $cfClass     = if ($jgErrors -gt 0) { 'company-flow cf-error' } elseif ($jgWarns -gt 0) { 'company-flow cf-warn' } else { 'company-flow' }

        # Company flow line
        $companyLineHtml = if ($companies.Found) {
            "<span class='$cfClass'><strong>$(ConvertTo-HtmlEncoded $companies.Source)</strong><span class='arrow'> &rarr; </span><strong>$(ConvertTo-HtmlEncoded $companies.Target)</strong></span>"
        } else {
            "<span class='$cfClass'>$(ConvertTo-HtmlEncoded $jobDesc)</span>"
        }

        # API status badge
        $apiBadgeHtml = if ($apiStatus) {
            $apiBadgeClass = switch ($apiStatus) {
                'Succeeded'          { 'api-badge api-ok'   }
                'PartiallySucceeded' { 'api-badge api-warn' }
                default              { 'api-badge api-err'  }
            }
            "&nbsp;&nbsp;<span class='$apiBadgeClass'>$(ConvertTo-HtmlEncoded $apiStatus)</span>"
        } else { '' }

        # Mini stats
        $statsHtml = "<span>$($groupItems.Count) entit$(if ($groupItems.Count -eq 1){'y'}else{'ies'})</span>"
        if ($jgErrors -gt 0) { $statsHtml += "&ensp;&bull;&ensp;<span class='stat-err'>$jgErrors error$(if ($jgErrors -eq 1){''}else{'s'})</span>" }
        if ($jgWarns  -gt 0) { $statsHtml += "&ensp;&bull;&ensp;<span class='stat-warn'>$jgWarns warning$(if ($jgWarns -eq 1){''}else{'s'})</span>" }
        if ($jgFailed -gt 0) { $statsHtml += "&ensp;&bull;&ensp;<span class='stat-err'>$jgFailed failed record$(if ($jgFailed -eq 1){''}else{'s'})</span>" }

        [void]$tbodySb.AppendLine("<tr class='$jobRowClass'>")
        [void]$tbodySb.AppendLine("  <td colspan='8'>")
        [void]$tbodySb.AppendLine("    <div class='job-top-line'>$companyLineHtml$apiBadgeHtml</div>")
        if ($companies.Found) {
            [void]$tbodySb.AppendLine("    <div class='job-desc'>$(ConvertTo-HtmlEncoded $jobDesc)</div>")
        }
        [void]$tbodySb.AppendLine("    <div class='job-meta-row'><span class='exec-id'>Execution ID: $(ConvertTo-HtmlEncoded $jg.Name)</span>&ensp;&bull;&ensp;$statsHtml</div>")
        [void]$tbodySb.AppendLine("  </td>")
        [void]$tbodySb.AppendLine("</tr>")

        # Entity rows for this job
        $entityRows = @(if ($IssuesOnly) {
            $groupItems | Where-Object { $_.NeedsAttention } | Sort-Object DefinitionGroupId, EntityName
        } else {
            $groupItems | Sort-Object DefinitionGroupId, EntityName
        })

        if ($entityRows.Count -eq 0) {
            [void]$tbodySb.AppendLine("<tr class='r-ok'><td colspan='8' class='no-issues-cell'>&#10003;&nbsp; All entities completed without issues</td></tr>")
        } else {
            foreach ($d in $entityRows) {
                $rowClass    = if ($d.IsError) { 'r-err' } elseif ($d.NeedsAttention) { 'r-warn' } else { 'r-ok' }
                $failedCellHtml = if ($d.FailedRecords -gt 0) {
                    "<td class='r fail-cell'>$($d.FailedRecords)</td>"
                } else {
                    "<td class='r'>0</td>"
                }

                [void]$tbodySb.AppendLine("<tr class='$rowClass'>")
                [void]$tbodySb.AppendLine("  <td>$(ConvertTo-HtmlEncoded $d.DefinitionGroupId)</td>")
                [void]$tbodySb.AppendLine("  <td>$(ConvertTo-HtmlEncoded $d.EntityName)</td>")
                [void]$tbodySb.AppendLine("  <td>$(Get-StatusBadgeHtml $d.StagingStatus)</td>")
                [void]$tbodySb.AppendLine("  <td>$(Get-StatusBadgeHtml $d.TargetStatus)</td>")
                [void]$tbodySb.AppendLine("  <td class='r'>$($d.StagingRecordsToBeProcessedCount)</td>")
                [void]$tbodySb.AppendLine("  <td class='r'>$($d.TargetRecordsCreatedCount)</td>")
                [void]$tbodySb.AppendLine("  <td class='r'>$($d.TargetRecordsUpdatedCount)</td>")
                [void]$tbodySb.AppendLine($failedCellHtml)
                [void]$tbodySb.AppendLine("</tr>")
            }
        }
    }

    $tbodyHtml = $tbodySb.ToString()

    # -- Assemble full HTML document ------------------------------------------
    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DMF Execution Job Report</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;font-size:13px;background:#f3f2f1;color:#323130;line-height:1.5}

/* ── Header ─────────────────────────────────────────── */
.page-header{background:#1b3a6b;color:#fff;padding:22px 40px}
.page-header h1{font-size:20px;font-weight:600;margin-bottom:6px}
.page-header .meta{font-size:12px;opacity:.8}
.page-header .meta span{margin-right:28px}

/* ── Summary cards ───────────────────────────────────── */
.summary{display:flex;gap:14px;padding:22px 40px 8px;flex-wrap:wrap}
.card{background:#fff;border-radius:3px;padding:14px 22px 12px;min-width:110px;box-shadow:0 1px 3px rgba(0,0,0,.09);border-top:3px solid #0078d4}
.card.ok   {border-top-color:#107c10}
.card.warn {border-top-color:#ca5010}
.card.err  {border-top-color:#d13438}
.card .val {font-size:30px;font-weight:700;line-height:1.1}
.card .lbl {font-size:11px;color:#797775;text-transform:uppercase;letter-spacing:.5px;margin-top:4px}
.card.ok   .val{color:#107c10}
.card.warn .val{color:#ca5010}
.card.err  .val{color:#d13438}

/* ── Filter notice ───────────────────────────────────── */
.filter-notice{margin:12px 40px 0;padding:9px 14px;background:#fff4ce;border-left:3px solid #f2c94c;font-size:12px;color:#605e5c;border-radius:0 3px 3px 0}

/* ── Table wrapper ───────────────────────────────────── */
.table-wrap{padding:20px 40px 40px}

table{width:100%;border-collapse:collapse;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.09);border-radius:3px;overflow:hidden}

/* Column headers */
thead th{background:#1b3a6b;color:#fff;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.4px;padding:10px 12px;text-align:left;white-space:nowrap;position:sticky;top:0;z-index:1}
thead th.r{text-align:right}

/* Job group rows */
tr.job-row td{background:#eef2f9;border-top:2px solid #1b3a6b;border-bottom:1px solid #c7d3e8;padding:10px 14px}
tr.job-row.has-errors   td{background:#fef0f0;border-top-color:#d13438;border-bottom-color:#f5c6c6}
tr.job-row.has-warnings td{background:#fef8f0;border-top-color:#ca5010;border-bottom-color:#f5d9c0}

.job-top-line{margin-bottom:3px}
.company-flow{font-size:15px;font-weight:700;color:#1b3a6b}
.company-flow .arrow{color:#0078d4;font-weight:400}
.cf-error{color:#d13438}
.cf-error .arrow{color:#d13438}
.cf-warn {color:#ca5010}
.cf-warn  .arrow{color:#ca5010}
.job-desc{font-size:12px;color:#605e5c;margin-bottom:4px}
.job-meta-row{font-size:11px;color:#797775;margin-top:3px}
.exec-id{font-family:'Consolas','Courier New',monospace;font-size:11px}
.stat-err {color:#d13438;font-weight:600}
.stat-warn{color:#ca5010;font-weight:600}

/* API status badges */
.api-badge{display:inline-block;padding:1px 9px;border-radius:10px;font-size:11px;font-weight:600}
.api-ok  {background:#dff6dd;color:#107c10}
.api-warn{background:#fff4ce;color:#7a4f00}
.api-err {background:#fde7e9;color:#c50f1f}

/* Entity rows */
tr.r-ok  {background:#fff}
tr.r-warn{background:#fffbf0}
tr.r-err {background:#fff5f5}
tr.r-ok:hover  {background:#faf9f8}
tr.r-warn:hover{background:#fff8e5}
tr.r-err:hover {background:#ffeded}

td{padding:7px 12px;border-bottom:1px solid #f0efee;vertical-align:middle}
td.r{text-align:right;font-variant-numeric:tabular-nums}
td.fail-cell{text-align:right;color:#d13438;font-weight:700}
td.no-issues-cell{color:#107c10;font-style:italic;padding:8px 14px}
tr:last-child td{border-bottom:none}

/* Status badges */
.badge{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;font-weight:600;white-space:nowrap}
.b-ok  {background:#dff6dd;color:#107c10}
.b-err {background:#fde7e9;color:#c50f1f}
.b-warn{background:#fff4ce;color:#7a4f00}
.b-none{background:#f3f2f1;color:#605e5c}

/* Footer */
footer{padding:14px 40px;font-size:11px;color:#a19f9d;border-top:1px solid #edebe9;background:#faf9f8}

@media print{
  body{background:#fff}
  .page-header,thead th,tr.job-row td{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  table{box-shadow:none;border:1px solid #edebe9}
  thead th{position:static}
}
</style>
</head>
<body>

<header class="page-header">
  <h1>D365 F&amp;O &mdash; Data Management Execution Job Report</h1>
  <div class="meta">
    <span>Environment: $(ConvertTo-HtmlEncoded $EnvironmentUrl)</span>
    <span>Generated: $reportTimestamp</span>
    <span>Jobs: $($matchedJobs.Count)</span>$(if ($IssuesOnly) { '<span>&#9888; Issues Only</span>' })
  </div>
</header>

<section class="summary">
  <div class="card">
    <div class="val">$($matchedJobs.Count)</div>
    <div class="lbl">Jobs</div>
  </div>
  <div class="card">
    <div class="val">$($allDetails.Count)</div>
    <div class="lbl">Total Entities</div>
  </div>
  <div class="card ok">
    <div class="val">$okCount</div>
    <div class="lbl">Clean</div>
  </div>
  <div class="card warn">
    <div class="val">$warnCount</div>
    <div class="lbl">Warnings</div>
  </div>
  <div class="card err">
    <div class="val">$errorCount</div>
    <div class="lbl">Errors</div>
  </div>
  <div class="card $(if ($totalFailed -gt 0) {'err'} else {'ok'})">
    <div class="val">$totalFailed</div>
    <div class="lbl">Failed Records</div>
  </div>
</section>

$filterNoticeHtml

<div class="table-wrap">
<table>
  <thead>
    <tr>
      <th>Data Package</th>
      <th>Entity</th>
      <th>Staging</th>
      <th>Target</th>
      <th class="r">Staged</th>
      <th class="r">Created</th>
      <th class="r">Updated</th>
      <th class="r">Failed</th>
    </tr>
  </thead>
  <tbody>
$tbodyHtml  </tbody>
</table>
</div>

<footer>
  D365 F&amp;O Data Management Execution Job Report &nbsp;&bull;&nbsp; EnvironmentDataLoader / Get-ExecutionJobReport.ps1 &nbsp;&bull;&nbsp; $reportTimestamp
</footer>

</body>
</html>
"@

    [System.IO.File]::WriteAllText($HtmlPath, $htmlContent, [System.Text.Encoding]::UTF8)
    Write-OK "HTML report written : $HtmlPath"
}

# =============================================================================
#  10. Finish
# =============================================================================
Stop-RunTranscript

if ($PassThru) {
    $allDetails | ForEach-Object { Write-Output $_ }
}

if ($errorCount -gt 0) { exit 1 }
