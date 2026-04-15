#Requires -Version 5.1
<#
.SYNOPSIS
    Uploads pre-built D365 F&O DMF package .zip files from a local directory
    and imports them into the specified legal entity.

.DESCRIPTION
    1. Scans UploadPath for .zip files, inspecting each zip's embedded
       Manifest.xml to surface the definition group name and entity count.
    2. Presents an interactive selection menu (skipped when -PackageName is
       supplied).
    3. Prompts for confirmation before proceeding (suppressed with -Force).
    4. Authenticates once via Microsoft Entra device code flow; the token is
       reused across all selected packages.
    5. For each confirmed package:
       a. Requests an Azure Blob Storage write URL from D365.
       b. Uploads the .zip to blob storage.
       c. Submits the import via ImportFromPackage against the specified
          legal entity.
       d. Polls until a terminal status is reached, showing a live progress bar.
    6. Prints a colour-coded summary table with per-package elapsed time and
       execution IDs for follow-up in D365 Job history.
    7. Writes a timestamped transcript to -LogPath (auto-generated when
       omitted; pass an empty string to suppress logging entirely).

    -WhatIf validates all selected packages (reads each zip locally) and shows
    what would be imported without making any API calls to D365.

    Unlike Invoke-BaselineImport.ps1, this script accepts packages that are
    already fully assembled .zip files (e.g. exports from another environment
    or downloads from Invoke-TemplateExport.ps1).  No manifest rebuilding or
    entity ordering is performed.

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
    D365 legal entity (company) to import into.
    Example: DAT

.PARAMETER UploadPath
    Directory containing pre-built DMF package .zip files to upload.

.PARAMETER PackageName
    When supplied, imports this one .zip file without prompting.
    The value may include or omit the .zip extension.
    Example: 'SystemSetupExport_20240101.zip'

.PARAMETER LogPath
    Path for the run transcript log.  When omitted, a log is auto-generated in
    UploadPath as DMFUpload_<timestamp>.log.  Pass an empty string ('') to
    suppress transcript logging entirely.

.PARAMETER PollIntervalSeconds
    How often (in seconds) to check import status.  Range: 5-300.  Default: 30.

.PARAMETER TimeoutMinutes
    Per-package polling timeout in minutes.  Range: 1-480.  Default: 60.

.PARAMETER MaxRetries
    Maximum number of automatic retries for transient REST failures.
    Range: 0-10.  Default: 3.

.PARAMETER Force
    Skip the confirmation prompt before importing.

.PARAMETER WhatIf
    Validate and display what would be imported without calling D365 APIs.
    No authentication is performed and no data is changed.

.PARAMETER NoOverwrite
    Preserve existing records in D365.  By default, existing records are
    overwritten during import.

.PARAMETER PassThru
    Return the per-package result objects to the pipeline after completion.
    Each object has properties: Package, DefinitionGroupId, Status,
    ExecutionId, Elapsed.

.EXAMPLE
    # Interactive: list all .zip files in C:\Exports and prompt for selection
    .\Invoke-PackageUpload.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'DAT' `
        -UploadPath     'C:\Exports'

.EXAMPLE
    # Non-interactive: upload one package, skip confirmation, preserve records
    .\Invoke-PackageUpload.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'DAT' `
        -UploadPath     'C:\Exports' `
        -PackageName    'SystemSetupExport_20240101.zip' `
        -NoOverwrite `
        -Force

.EXAMPLE
    # Dry-run: validate all packages in the directory without making API calls
    .\Invoke-PackageUpload.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'DAT' `
        -UploadPath     'C:\Exports' `
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

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$UploadPath,

    [string]$PackageName,

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
    [switch]$NoOverwrite,
    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression.FileSystem

# =============================================================================
#  Load library modules
# =============================================================================
$libPath = Join-Path $PSScriptRoot 'lib'
. (Join-Path $libPath 'DmfOutput.ps1')
. (Join-Path $libPath 'DmfRequest.ps1')
. (Join-Path $libPath 'DmfZip.ps1')

# =============================================================================
#  Pre-flight path validation  (before transcript so errors surface cleanly)
# =============================================================================
if (-not (Test-Path $UploadPath -PathType Container)) {
    throw "Upload path not found: '$UploadPath'"
}

# =============================================================================
#  Script-level constants  (consumed by lib functions via $Script: scope)
# =============================================================================
$Script:Version          = '1.0'
$Script:DmNs             = 'http://schemas.microsoft.com/dynamics/2015/01/DataManagement'
$Script:LineWidth        = 80
$Script:MaxRetries       = $MaxRetries
$Script:TranscriptActive = $false

try {
    $w = $Host.UI.RawUI.BufferSize.Width
    $Script:LineWidth = [Math]::Min(120, [Math]::Max(80, $w))
} catch { <# non-interactive host -- keep default #> }

# =============================================================================
#  1.  Transcript startup  (UploadPath validated above)
# =============================================================================
if (-not $PSBoundParameters.ContainsKey('LogPath')) {
    $ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
    $LogPath = Join-Path $UploadPath "DMFUpload_${ts}.log"
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
Write-Banner -DryRun:$WhatIf -Title 'D365 F&O Package Upload Utility'
Write-Host ''
Write-Info "Environment  : $EnvironmentUrl"
Write-Info "Tenant       : $TenantId"
Write-Info "Legal entity : $LegalEntityId"
Write-Info "Upload from  : $UploadPath"
if ($Script:TranscriptActive) { Write-Info "Log          : $LogPath" }
if ($WhatIf)      { Write-Warn 'WhatIf active -- no API calls will be made.' }
if ($Force)       { Write-Info 'Force        : confirmation prompt suppressed' }
if ($NoOverwrite) { Write-Info 'NoOverwrite  : existing records will not be overwritten' }

# =============================================================================
#  3.  Discover packages
# =============================================================================
Write-Step 'Discovering packages'

$allZips = @(Get-ChildItem -Path $UploadPath -Filter '*.zip' -File)
if ($allZips.Count -eq 0) {
    throw "No .zip files found under '$UploadPath'."
}

Write-Info "Scanning $($allZips.Count) zip file(s)..."
$allPackages = @(for ($i = 0; $i -lt $allZips.Count; $i++) {
    Get-ZipPackageInfo -File $allZips[$i] -Index ($i + 1)
})

# =============================================================================
#  4.  Package selection
# =============================================================================
$selectedPackages = [System.Collections.Generic.List[pscustomobject]]::new()

if ($PackageName) {
    # ── Single-package mode (non-interactive) ──────────────────────────────
    $needle = $PackageName
    if (-not $needle.EndsWith('.zip', [System.StringComparison]::OrdinalIgnoreCase)) {
        $needle = "$needle.zip"
    }
    $match = $allPackages | Where-Object { $_.Name -ieq $needle }
    if (-not $match) {
        $available = ($allPackages | ForEach-Object { "    '$($_.Name)'" }) -join [System.Environment]::NewLine
        throw "Package '$PackageName' not found under '$UploadPath'.`nAvailable packages:`n$available"
    }
    $selectedPackages.Add($match)
    Write-Info "Package : $($match.Name)"
} else {
    # ── Interactive numbered selection ─────────────────────────────────────
    Write-Step 'Select packages to import'

    $nameWidth = ($allPackages | ForEach-Object { $_.Name.Length } | Measure-Object -Maximum).Maximum
    $nameWidth = [int][Math]::Max(30, [Math]::Min($nameWidth, $Script:LineWidth - 40))

    $colFmt = '  {0,3}  {1}  {2,8}  {3,8}  {4}'
    Write-Host ''
    Write-Host ($colFmt -f '#', 'Package'.PadRight($nameWidth), 'Size(MB)', 'Entities', 'Definition Group') -ForegroundColor White
    Write-Host ($colFmt -f '---', ('-' * $nameWidth), '--------', '--------', '----------------') -ForegroundColor DarkGray

    foreach ($pkg in $allPackages) {
        $nameCol    = if ($pkg.Name.Length -gt $nameWidth) {
                          $pkg.Name.Substring(0, $nameWidth - 3) + '...'
                      } else { $pkg.Name.PadRight($nameWidth) }
        $sizeCol    = $pkg.SizeMB.ToString('F2').PadLeft(8)
        $entCol     = if ($pkg.HasManifest) { $pkg.EntityCount.ToString().PadLeft(8) } else { '       !' }
        $dgCol      = if ($pkg.DefinitionGroupId) { $pkg.DefinitionGroupId } else { '(unknown)' }
        $color      = if ($pkg.IsValid) { 'White' } else { 'DarkYellow' }
        Write-Host ($colFmt -f $pkg.Index, $nameCol, $sizeCol, $entCol, $dgCol) -ForegroundColor $color
    }

    $invalidPkgs = @($allPackages | Where-Object { -not $_.IsValid })
    if ($invalidPkgs.Count -gt 0) {
        Write-Host ''
        foreach ($p in $invalidPkgs) {
            Write-Warn "  [!] '$($p.Name)': $($p.Warnings -join '; ')"
        }
    }

    # Validated input loop
    $selectedIndices = $null
    do {
        Write-Host ''
        $rawInput = (Read-Host '  Selection  (e.g. 1   1,3   2-4   A=all   Q=quit)').Trim().ToLower()

        if ($rawInput -in 'q', 'quit') { $selectedIndices = @(); break }
        if ($rawInput -in 'a', 'all', '') { $selectedIndices = 1..$allPackages.Count; break }

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

        $outOfRange = @($parsed | Where-Object { $_ -lt 1 -or $_ -gt $allPackages.Count })
        if ($outOfRange.Count -gt 0) {
            Write-Warn "  Out-of-range: $($outOfRange -join ', ').  Valid: 1-$($allPackages.Count)."
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
        Write-Warn 'No packages selected.  Exiting.'
        Stop-RunTranscript
        exit 0
    }

    foreach ($idx in $selectedIndices) {
        $selectedPackages.Add(($allPackages | Where-Object { $_.Index -eq $idx }))
    }
}

# =============================================================================
#  5.  WhatIf exit
# =============================================================================
if ($WhatIf) {
    Write-Host ''
    Write-Rule 'WhatIf -- validation complete, no import will occur'
    Write-Host ''
    Write-Info "Environment  : $EnvironmentUrl"
    Write-Info "Legal entity : $LegalEntityId"
    Write-Info "Packages     : $($selectedPackages.Count)"
    foreach ($pkg in $selectedPackages) {
        $meta = "$($pkg.SizeMB) MB"
        if ($pkg.HasManifest) { $meta += ", $($pkg.EntityCount) entities" }
        if ($pkg.DefinitionGroupId) { $meta += ", def group: $($pkg.DefinitionGroupId)" }
        if (-not $pkg.IsValid) { $meta += "  [!] $($pkg.Warnings -join '; ')" }
        Write-Detail "[$($pkg.Index)] $($pkg.Name)  ($meta)"
    }
    Write-Host ''
    Write-Info 'No API calls made.  Remove -WhatIf to perform the import.'
    if ($PassThru) {
        $selectedPackages | ForEach-Object {
            [pscustomobject]@{
                Package         = $_.Name
                DefinitionGroupId = $_.DefinitionGroupId
                Status          = 'WhatIf'
                ExecutionId     = '-'
                Elapsed         = '-'
            }
        }
    }
    Stop-RunTranscript
    exit 0
}

# =============================================================================
#  6.  Confirmation  (skipped with -Force)
# =============================================================================
Write-Host ''
Write-Rule 'Ready to import'
Write-Host ''
Write-Info "Environment  : $EnvironmentUrl"
Write-Info "Legal entity : $LegalEntityId"
Write-Info "Overwrite    : $(if ($NoOverwrite) { 'No -- existing records will be preserved' } else { 'Yes -- existing records will be replaced' })"
Write-Info "Packages     : $($selectedPackages.Count)"

foreach ($pkg in $selectedPackages) {
    $meta = "$($pkg.SizeMB) MB"
    if ($pkg.HasManifest) { $meta += ", $($pkg.EntityCount) entities" }
    if ($pkg.DefinitionGroupId) { $meta += ", def group: $($pkg.DefinitionGroupId)" }
    if (-not $pkg.IsValid) { $meta += "  [!] $($pkg.Warnings -join '; ')" }
    Write-Detail "[$($pkg.Index)] $($pkg.Name)  ($meta)"
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
#  7.  Authenticate  (once -- token reused across all packages)
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

# =============================================================================
#  8.  Process packages
# =============================================================================
$results     = [System.Collections.Generic.List[pscustomobject]]::new()
$scriptStart = Get-Date

foreach ($pkg in $selectedPackages) {
    $pkgName    = $pkg.Name
    $pkgSafe    = ($pkgName -replace '\.zip$', '' -replace '[^A-Za-z0-9]', '-') -replace '-{2,}', '-'
    $pkgStart   = Get-Date
    $timestamp  = Get-Date -Format 'yyyyMMddHHmmss'
    $pkgResult  = [pscustomobject]@{
        Package           = $pkgName
        DefinitionGroupId = $pkg.DefinitionGroupId
        Status            = 'Error'
        ExecutionId       = '-'
        Elapsed           = '-'
    }

    $divider = '=' * $Script:LineWidth
    Write-Host ''
    Write-Host $divider -ForegroundColor DarkCyan
    Write-Host "  [$($results.Count + 1)/$($selectedPackages.Count)]  $pkgName" -ForegroundColor Cyan
    if ($pkg.DefinitionGroupId) {
        Write-Host "  Definition group: $($pkg.DefinitionGroupId)" -ForegroundColor DarkGray
    }
    Write-Host $divider -ForegroundColor DarkCyan

    # Token expiry warnings
    if ((Get-Date) -ge $tokenExpiry) {
        Write-Warn 'Access token has expired.  API calls will likely fail with HTTP 401.  Re-run the script.'
    } elseif ((Get-Date).AddMinutes(5) -ge $tokenExpiry) {
        Write-Warn "Token expires at $($tokenExpiry.ToString('HH:mm:ss')) -- it may expire mid-import."
    }

    try {
        # Guard: refuse to proceed if the zip failed validation
        if (-not $pkg.IsValid) {
            throw "Package validation failed: $($pkg.Warnings -join '; ')"
        }

        $definitionGroupId = $pkg.DefinitionGroupId
        $zipPath           = $pkg.File.FullName
        $zipSizeMB         = $pkg.SizeMB

        # ── a. Request Azure Blob upload URL ──────────────────────────────────
        Write-Step 'Requesting blob upload URL'
        $blobFileName = "${pkgSafe}_${timestamp}.zip"
        $writeResp = Invoke-DmfRequest -Operation 'GetAzureWriteUrl' -Params @{
            Method      = 'Post'
            Uri         = "$dmfBase.GetAzureWriteUrl"
            Headers     = $authHeaders
            ContentType = 'application/json'
            Body        = (@{ uniqueFileName = $blobFileName } | ConvertTo-Json)
        }

        # Response is Edm.String -- the payload is a JSON-encoded string inside 'value'
        $blobResult = $writeResp.value | ConvertFrom-Json
        $blobUrl    = $blobResult.BlobUrl
        $blobId     = $blobResult.BlobId
        Write-Info "Blob ID : $blobId"

        if ([string]::IsNullOrWhiteSpace($blobUrl)) {
            throw "GetAzureWriteUrl returned an empty BlobUrl.  Response: $($writeResp | ConvertTo-Json -Depth 5)"
        }

        # ── b. Upload zip to Azure Blob ───────────────────────────────────────
        Write-Step "Uploading package ($zipSizeMB MB)"
        $null = Invoke-DmfRequest -Operation 'blob upload' -Params @{
            Method      = 'Put'
            Uri         = $blobUrl
            Headers     = @{ 'x-ms-blob-type' = 'BlockBlob' }
            ContentType = 'application/octet-stream'
            InFile      = $zipPath
        }
        Write-Info 'Upload complete.'

        # ── c. Submit import job ──────────────────────────────────────────────
        Write-Step "Submitting import job  (legal entity: $LegalEntityId)"
        Write-Info "Definition group : $definitionGroupId"
        $importResp = Invoke-DmfRequest -Operation 'ImportFromPackage' -Params @{
            Method      = 'Post'
            Uri         = "$dmfBase.ImportFromPackage"
            Headers     = $authHeaders
            ContentType = 'application/json'
            Body        = (@{
                packageUrl        = $blobUrl
                definitionGroupId = $definitionGroupId
                executionId       = ''
                execute           = $true
                overwrite         = (-not $NoOverwrite.IsPresent)
                legalEntityId     = $LegalEntityId
            } | ConvertTo-Json)
        }

        $executionId = $importResp.value
        if ([string]::IsNullOrWhiteSpace($executionId)) {
            throw "ImportFromPackage returned an empty execution ID.  Response: $($importResp | ConvertTo-Json -Depth 5)"
        }
        Write-Info "Execution ID : $executionId"

        # ── d. Poll for completion (live progress bar) ────────────────────────
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
                    -Activity    "Importing: $pkgName" `
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
            Write-Progress -Id 1 -Activity "Importing: $pkgName" -Completed
        }

        # ── e. Package outcome ────────────────────────────────────────────────
        $pkgElapsed            = (Get-Date) - $pkgStart
        $pkgResult.ExecutionId = $executionId
        $pkgResult.Elapsed     = Format-Elapsed $pkgElapsed
        $pkgResult.Status      = if ($lastStatus) { $lastStatus } else { 'TimedOut' }
        $jobPath               = "Data management > Job history > execution ID: $executionId"

        switch ($lastStatus) {
            'Succeeded' {
                Write-OK "Package imported successfully.  ($($pkgResult.Elapsed))"
            }
            'PartiallySucceeded' {
                Write-Warn 'Import partially succeeded -- some records may have been skipped or errored.'
                Write-Info "Review : $jobPath"
            }
            'Failed' {
                Write-Fail 'Import failed.'
                Write-Info "Review : $jobPath"
            }
            'Canceled' {
                Write-Fail 'Import was canceled.'
                Write-Info "Execution ID : $executionId"
            }
            default {
                Write-Warn "Timed out after $TimeoutMinutes minutes.  Last known status: '$lastStatus'"
                Write-Info "Execution ID : $executionId"
            }
        }
    }
    catch {
        $pkgResult.Elapsed = Format-Elapsed ((Get-Date) - $pkgStart)
        $pkgResult.Status  = 'Error'
        Write-Fail "Package '$pkgName' failed:"
        Write-Fail "  $_"
        Write-Verbose $_.ScriptStackTrace
    }

    $results.Add($pkgResult)
    if ($PassThru) { Write-Output $pkgResult }
}

# =============================================================================
#  9.  Summary
# =============================================================================
$totalElapsed = Format-Elapsed ((Get-Date) - $scriptStart)
$divider      = '=' * $Script:LineWidth

Write-Host ''
Write-Host $divider -ForegroundColor DarkCyan
Write-Host "  Summary  --  $($results.Count) package(s)  |  total time: $totalElapsed" -ForegroundColor Cyan
Write-Host $divider -ForegroundColor DarkCyan
Write-Host ''

$nameWidth2   = ($results | ForEach-Object { $_.Package.Length } | Measure-Object -Maximum).Maximum
$nameWidth2   = [Math]::Max(20, [Math]::Min($nameWidth2, $Script:LineWidth - 55))
$statusWidth  = 20
$elapsedWidth = 10

$summaryHdr = '  {0}  {1}  {2}  {3}' -f 'Status'.PadRight($statusWidth), 'Elapsed'.PadRight($elapsedWidth), 'Package'.PadRight($nameWidth2), 'Execution ID'
$summarySep = '  {0}  {1}  {2}  {3}' -f ('-' * $statusWidth), ('-' * $elapsedWidth), ('-' * $nameWidth2), ('-' * 38)

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

    $nameTrunc = if ($r.Package.Length -gt $nameWidth2) {
        $r.Package.Substring(0, $nameWidth2 - 3) + '...'
    } else { $r.Package.PadRight($nameWidth2) }

    $row = '  {0}  {1}  {2}  {3}' -f $r.Status.PadRight($statusWidth), $r.Elapsed.PadRight($elapsedWidth), $nameTrunc, $r.ExecutionId
    Write-Host $row -ForegroundColor $color
}

Write-Host ''
Write-Host $divider -ForegroundColor DarkCyan

$resultColor = if ($failCount -eq 0) { 'Green' } else { 'Red' }
$resultMsg   = if ($failCount -eq 0) {
    "  All $successCount package(s) completed successfully."
} else {
    "  $successCount succeeded  |  $failCount failed."
}
Write-Host $resultMsg -ForegroundColor $resultColor
Write-Host $divider   -ForegroundColor DarkCyan
Write-Host ''

Stop-RunTranscript

if ($failCount -gt 0) { exit 1 }
