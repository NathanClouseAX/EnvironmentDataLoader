#Requires -Version 5.1
<#
.SYNOPSIS
    Discovers D365 F&O data packages under the resources folder and imports them
    via the Data Management Framework API.

.DESCRIPTION
    1. Scans all subfolders under ResourcesPath, treating each as an importable package.
    2. Presents an interactive selection menu (skipped when -PackageName is supplied).
    3. Prompts for confirmation before authenticating.
    4. Authenticates once via Microsoft Entra device code flow; the token is reused
       across all selected packages.
    5. For each confirmed package:
       a. Discovers .xlsx files on disk and cross-references them against Manifest.xml.
       b. Rebuilds the manifest with optimised execution-unit / level / sequence ordering
          so that independent entity chains run in parallel.
       c. Builds a well-formed .zip package and uploads it to Azure Blob Storage.
       d. Submits the import via ImportFromPackage and polls until a terminal status
          is reached, showing a live progress bar throughout.
    6. Prints a colour-coded summary table with per-package elapsed time and
       execution IDs for follow-up in D365 Job history.

    All REST calls include automatic retry with exponential back-off (HTTP 5xx /
    429 / transient network errors).  The token expiry window is tracked and
    surfaced as a warning when approaching expiry.

.PARAMETER EnvironmentUrl
    Base URL of the D365 F&O environment (no trailing slash).
    Example: https://contoso.operations.dynamics.com

.PARAMETER TenantId
    Microsoft Entra tenant ID (GUID) or verified domain name.
    Example: contoso.onmicrosoft.com

.PARAMETER LegalEntityId
    D365 legal entity to import into.
    Example: DAT

.PARAMETER PackageName
    When supplied, imports this one package folder without prompting.
    The value must match a subfolder name exactly.
    Example: 'Demo data 10.0.9 100 System and Shared'

.PARAMETER ResourcesPath
    Root folder that contains package subfolders.
    Defaults to the repo-relative 'resources' directory next to this script.

.PARAMETER OutputPath
    Directory for generated .zip files.  Defaults to $env:TEMP.

.PARAMETER PollIntervalSeconds
    How often (in seconds) to check import status.  Range: 5-300.  Default: 30.

.PARAMETER TimeoutMinutes
    Per-package polling timeout in minutes.  Range: 1-480.  Default: 60.

.PARAMETER MaxRetries
    Maximum number of automatic retries for transient REST failures.
    Range: 0-10.  Default: 3.

.EXAMPLE
    # Interactive: discover all packages and prompt for each
    .\Invoke-BaselineImport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'DAT'

.EXAMPLE
    # Non-interactive: import one specific package
    .\Invoke-BaselineImport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'DAT' `
        -PackageName    'Demo data 10.0.9 100 System and Shared'
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

    [string]$PackageName,

    [ValidateNotNullOrEmpty()]
    [string]$ResourcesPath = (Join-Path $PSScriptRoot 'resources'),

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = $env:TEMP,

    [ValidateRange(5, 300)]
    [int]$PollIntervalSeconds = 30,

    [ValidateRange(1, 480)]
    [int]$TimeoutMinutes = 60,

    [ValidateRange(0, 10)]
    [int]$MaxRetries = 3
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression.FileSystem

# =============================================================================
#  Script-level constants
# =============================================================================
$Script:Version    = '2.0'
$Script:DmNs       = 'http://schemas.microsoft.com/dynamics/2015/01/DataManagement'
$Script:LineWidth  = 80
$Script:MaxRetries = $MaxRetries
try {
    $w = $Host.UI.RawUI.BufferSize.Width
    $Script:LineWidth = [Math]::Min(120, [Math]::Max(80, $w))
} catch { <# non-interactive host -- keep default #> }

# =============================================================================
#  Output helpers
# =============================================================================
function Write-Banner {
    $w    = $Script:LineWidth
    $line = '=' * $w
    Write-Host $line                                           -ForegroundColor DarkCyan
    Write-Host "  D365 F&O Data Import Utility  v$Script:Version" -ForegroundColor Cyan
    Write-Host $line                                           -ForegroundColor DarkCyan
}

function Write-Rule {
    param([string]$Label = '')
    $w = $Script:LineWidth
    if ($Label) {
        $pad = [Math]::Max(1, $w - $Label.Length - 5)
        Write-Host "--- $Label $('-' * $pad)" -ForegroundColor DarkGray
    } else {
        Write-Host ('-' * $w) -ForegroundColor DarkGray
    }
}

function Write-Step   { param([string]$m) Write-Host "`n==> $m"   -ForegroundColor Cyan    }
function Write-Info   { param([string]$m) Write-Host "    $m"     -ForegroundColor Gray    }
function Write-Detail { param([string]$m) Write-Host "      $m"   -ForegroundColor DarkGray }
function Write-OK     { param([string]$m) Write-Host "`n[OK]   $m" -ForegroundColor Green   }
function Write-Warn   { param([string]$m) Write-Host "[WARN] $m"  -ForegroundColor Yellow  }
function Write-Fail   { param([string]$m) Write-Host "[FAIL] $m"  -ForegroundColor Red     }

function Format-Elapsed {
    param([TimeSpan]$Elapsed)
    $h = [int][Math]::Floor($Elapsed.TotalHours)
    $m = $Elapsed.Minutes
    $s = $Elapsed.Seconds
    if ($h -gt 0) { return "${h}h ${m}m ${s}s" }
    if ($m -gt 0) { return "${m}m ${s}s" }
    if ($s -gt 0) { return "${s}s" }
    return '<1s'
}

# =============================================================================
#  REST helper  --  automatic retry with exponential back-off
# =============================================================================
function Invoke-DmfRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable]$Params,
        [string]$Operation = 'API request',
        [int]$MaxRetries   = $Script:MaxRetries
    )

    Write-Verbose "[$Operation] $($Params.Method) $($Params.Uri)"

    for ($attempt = 1; $attempt -le ($MaxRetries + 1); $attempt++) {
        try {
            return Invoke-RestMethod @Params
        }
        catch {
            # -- Classify error -----------------------------------------------
            $httpStatus = 0
            $retryAfter = 0
            if ($null -ne $_.Exception.Response) {
                try { $httpStatus = [int]$_.Exception.Response.StatusCode } catch {}
                try {
                    $ra = $_.Exception.Response.Headers['Retry-After']
                    if ($ra) { $retryAfter = [int]$ra }
                } catch {}
            }

            # -- Extract OData / D365 error detail ----------------------------
            $detail = ''
            try {
                $errBody = $_.ErrorDetails.Message | ConvertFrom-Json
                if ($errBody.error.message)                    { $detail = $errBody.error.message }
                elseif ($errBody.'odata.error'.message.value)  { $detail = $errBody.'odata.error'.message.value }
            } catch {}
            if (-not $detail) { $detail = $_.Exception.Message }

            # -- Non-retryable client errors ----------------------------------
            if ($httpStatus -eq 401) {
                throw "[$Operation] HTTP 401 Unauthorized. The access token has expired or lacks permission. Re-run to re-authenticate."
            }
            if ($httpStatus -ge 400 -and $httpStatus -lt 500 -and $httpStatus -notin @(408, 429)) {
                throw "[$Operation] HTTP ${httpStatus}: $detail"
            }

            # -- Exhausted retries --------------------------------------------
            if ($attempt -gt $MaxRetries) {
                $label = if ($httpStatus -gt 0) { "HTTP $httpStatus" } else { 'network error' }
                throw "[$Operation] Failed after $MaxRetries retries ($label): $detail"
            }

            # -- Retryable: sleep then loop -----------------------------------
            $delay = if ($retryAfter -gt 0) { $retryAfter } else {
                [Math]::Min(30, [Math]::Pow(2, $attempt - 1) * 2)
            }
            $label = if ($httpStatus -gt 0) { "HTTP $httpStatus" } else { 'network error' }
            Write-Warn "[$Operation] $label - retrying in ${delay}s (attempt $attempt of $MaxRetries)..."
            Start-Sleep -Seconds $delay
        }
    }
}

# =============================================================================
#  Package discovery helper
# =============================================================================
function Get-PackageInfo {
    param(
        [Parameter(Mandatory)] [System.IO.DirectoryInfo]$Folder,
        [Parameter(Mandatory)] [int]$Index
    )

    $xlsxCount   = 0
    $entityCount = 0
    $hasManifest = $false
    $warnings    = [System.Collections.Generic.List[string]]::new()

    try {
        $xlsxCount = @(Get-ChildItem -Path $Folder.FullName -Filter '*.xlsx' -File -ErrorAction Stop).Count
    } catch { $warnings.Add('Could not enumerate .xlsx files') }

    $manifestPath = Join-Path $Folder.FullName 'Manifest.xml'
    if (Test-Path $manifestPath) {
        $hasManifest = $true
        try {
            $doc = New-Object System.Xml.XmlDocument
            $doc.Load($manifestPath)
            $entityCount = $doc.SelectNodes('//*[local-name()="DataManagementPackageEntityData"]').Count
        } catch { $warnings.Add('Manifest.xml could not be parsed') }
    } else {
        $warnings.Add('No Manifest.xml found')
    }

    return [pscustomobject]@{
        Index       = $Index
        Folder      = $Folder
        Name        = $Folder.Name
        XlsxCount   = $xlsxCount
        EntityCount = $entityCount
        HasManifest = $hasManifest
        IsValid     = ($hasManifest -and $xlsxCount -gt 0)
        Warnings    = $warnings
    }
}

# =============================================================================
#  Entity execution ordering
#
#  Entities sharing the same EU + LV + SEQ are processed in parallel.
#  Within an EU, levels run sequentially (ascending).
#  Within a level, sequences run sequentially (ascending).
#
#  EU=1  LV=10  Foundation + geographic + org data (prerequisite for all below)
#  EU=1  LV=20  Finance (GL) and Workflow merged into one level so their
#               independent chains advance in parallel at each matching SEQ step.
#
#  Entities absent from this table fall back silently to source Manifest.xml values.
# =============================================================================
$entityOrdering = [ordered]@{
    # -- EU=1  LV=10  SEQ=10 : Independent reference tables ------------------
    'Global address book parameters'                  = @{ EU=1; LV=10; SEQ=10 }
    'Language codes'                                  = @{ EU=1; LV=10; SEQ=10 }
    'Name affixes'                                    = @{ EU=1; LV=10; SEQ=10 }
    'Currencies'                                      = @{ EU=1; LV=10; SEQ=10 }
    'Ethnic origins'                                  = @{ EU=1; LV=10; SEQ=10 }
    'Veteran status'                                  = @{ EU=1; LV=10; SEQ=10 }
    'Name sequences'                                  = @{ EU=1; LV=10; SEQ=10 }
    'Relationship types'                              = @{ EU=1; LV=10; SEQ=10 }
    'Skill types'                                     = @{ EU=1; LV=10; SEQ=10 }
    'Titles'                                          = @{ EU=1; LV=10; SEQ=10 }
    'User groups'                                     = @{ EU=1; LV=10; SEQ=10 }
    'Rating models'                                   = @{ EU=1; LV=10; SEQ=10 }
    'Regulatory establishments'                       = @{ EU=1; LV=10; SEQ=10 }
    'Address format'                                  = @{ EU=1; LV=10; SEQ=10 }
    'Address books'                                   = @{ EU=1; LV=10; SEQ=10 }
    'Address and contact information purpose'         = @{ EU=1; LV=10; SEQ=10 }
    'Address parameters'                              = @{ EU=1; LV=10; SEQ=10 }
    'Organization hierarchy type'                     = @{ EU=1; LV=10; SEQ=10 }
    'Country/regions'                                 = @{ EU=1; LV=10; SEQ=10 }
    'Team types'                                      = @{ EU=1; LV=10; SEQ=10 }
    'System parameters'                               = @{ EU=1; LV=10; SEQ=10 }

    # -- EU=1  LV=10  SEQ=20 : Depends on SEQ=10 ----------------------------
    'Exchange rates'                                  = @{ EU=1; LV=10; SEQ=20 }
    'Rating level'                                    = @{ EU=1; LV=10; SEQ=20 }
    'Organization hierarchy purposes'                 = @{ EU=1; LV=10; SEQ=20 }
    'Address format lines'                            = @{ EU=1; LV=10; SEQ=20 }
    'Skills'                                          = @{ EU=1; LV=10; SEQ=20 }
    'States'                                          = @{ EU=1; LV=10; SEQ=20 }
    'Units'                                           = @{ EU=1; LV=10; SEQ=20 }
    'Legal entities'                                  = @{ EU=1; LV=10; SEQ=20 }

    # -- EU=1  LV=10  SEQ=30 : Depends on SEQ=20 ----------------------------
    'Counties'                                        = @{ EU=1; LV=10; SEQ=30 }
    'Unit translations'                               = @{ EU=1; LV=10; SEQ=30 }
    'Unit conversions'                                = @{ EU=1; LV=10; SEQ=30 }
    'Cities'                                          = @{ EU=1; LV=10; SEQ=30 }
    'Operating unit'                                  = @{ EU=1; LV=10; SEQ=30 }

    # -- EU=1  LV=10  SEQ=40 : Depends on SEQ=30 ----------------------------
    'Districts V2'                                    = @{ EU=1; LV=10; SEQ=40 }

    # -- EU=1  LV=10  SEQ=50 : Depends on SEQ=40 ----------------------------
    'Postal codes V3'                                 = @{ EU=1; LV=10; SEQ=50 }
    'Global address book V2'                          = @{ EU=1; LV=10; SEQ=50 }

    # -- EU=1  LV=10  SEQ=60 : Depends on SEQ=50 ----------------------------
    'Organization hierarchy V2 - published and draft' = @{ EU=1; LV=10; SEQ=60 }
    'User information'                                = @{ EU=1; LV=10; SEQ=60 }

    # -- EU=1  LV=10  SEQ=70 : Depends on SEQ=60 ----------------------------
    'User to person relationship'                     = @{ EU=1; LV=10; SEQ=70 }
    'Teams V2'                                        = @{ EU=1; LV=10; SEQ=70 }

    # -- EU=1  LV=10  SEQ=80 : Depends on SEQ=70 ----------------------------
    'Security user role association'                  = @{ EU=1; LV=10; SEQ=80 }

    # -- EU=1  LV=10  SEQ=90 : Depends on SEQ=80 ----------------------------
    'Party relationships'                             = @{ EU=1; LV=10; SEQ=90 }

    # -- EU=1  LV=10  SEQ=100 : Depends on SEQ=90 ---------------------------
    'Party contacts'                                  = @{ EU=1; LV=10; SEQ=100 }

    # -- EU=1  LV=10  SEQ=110 : Depends on SEQ=100 --------------------------
    'Party postal address V2'                         = @{ EU=1; LV=10; SEQ=110 }

    # -- EU=1  LV=20  SEQ=10 : Finance foundation + Workflow foundation ------
    'Chart of accounts'                               = @{ EU=1; LV=20; SEQ=10 }
    'Fiscal calendar'                                 = @{ EU=1; LV=20; SEQ=10 }
    'Financial dimensions'                            = @{ EU=1; LV=20; SEQ=10 }
    'Main account categories'                         = @{ EU=1; LV=20; SEQ=10 }
    'Expression'                                      = @{ EU=1; LV=20; SEQ=10 }
    'System email template'                           = @{ EU=1; LV=20; SEQ=10 }

    # -- EU=1  LV=20  SEQ=20 ------------------------------------------------
    'Main account'                                    = @{ EU=1; LV=20; SEQ=20 }
    'Fiscal calendar period'                          = @{ EU=1; LV=20; SEQ=20 }
    'Dimension attribute activation'                  = @{ EU=1; LV=20; SEQ=20 }
    'Financial dimension format'                      = @{ EU=1; LV=20; SEQ=20 }
    'Financial dimension translations'                = @{ EU=1; LV=20; SEQ=20 }
    'Workflow version'                                = @{ EU=1; LV=20; SEQ=20 }
    'Workflow parallel branch'                        = @{ EU=1; LV=20; SEQ=20 }
    'System email template message'                   = @{ EU=1; LV=20; SEQ=20 }

    # -- EU=1  LV=20  SEQ=30 ------------------------------------------------
    'Advanced rule structures'                        = @{ EU=1; LV=20; SEQ=30 }
    'Consolidation groups and accounts'               = @{ EU=1; LV=20; SEQ=30 }
    'Financial dimension values'                      = @{ EU=1; LV=20; SEQ=30 }
    'Financial dimension value translations'          = @{ EU=1; LV=20; SEQ=30 }
    'Financial dimension sets'                        = @{ EU=1; LV=20; SEQ=30 }
    'Workflow version notes'                          = @{ EU=1; LV=20; SEQ=30 }
    'Workflow subworkflow'                            = @{ EU=1; LV=20; SEQ=30 }
    'Workflow system parameters'                      = @{ EU=1; LV=20; SEQ=30 }
    'Workflow element'                                = @{ EU=1; LV=20; SEQ=30 }

    # -- EU=1  LV=20  SEQ=32 ------------------------------------------------
    'Advanced rule structure allowed values'          = @{ EU=1; LV=20; SEQ=32 }

    # -- EU=1  LV=20  SEQ=34 ------------------------------------------------
    'Advanced rule structure activation'              = @{ EU=1; LV=20; SEQ=34 }

    # -- EU=1  LV=20  SEQ=36 : Account structures + Workflow element details -
    'Account structures'                              = @{ EU=1; LV=20; SEQ=36 }
    'Workflow element action'                         = @{ EU=1; LV=20; SEQ=36 }
    'Workflow step'                                   = @{ EU=1; LV=20; SEQ=36 }
    'Workflow element notification'                   = @{ EU=1; LV=20; SEQ=36 }
    'Workflow element link'                           = @{ EU=1; LV=20; SEQ=36 }
    'Workflow element outcome message'                = @{ EU=1; LV=20; SEQ=36 }
    'Workflow version notification'                   = @{ EU=1; LV=20; SEQ=36 }

    # -- EU=1  LV=20  SEQ=38 ------------------------------------------------
    'Account structure allowed values'                = @{ EU=1; LV=20; SEQ=38 }

    # -- EU=1  LV=20  SEQ=40 : Advanced rules + Workflow notifications -------
    'Advanced rules'                                  = @{ EU=1; LV=20; SEQ=40 }
    'Number sequence code'                            = @{ EU=1; LV=20; SEQ=40 }
    'Workflow version notification message'           = @{ EU=1; LV=20; SEQ=40 }
    'Workflow escalation path'                        = @{ EU=1; LV=20; SEQ=40 }
    'Workflow element notification message'           = @{ EU=1; LV=20; SEQ=40 }
    'Workflow step message'                           = @{ EU=1; LV=20; SEQ=40 }
    'Workflow line item'                              = @{ EU=1; LV=20; SEQ=40 }

    # -- EU=1  LV=20  SEQ=42 ------------------------------------------------
    'Advanced rule criteria'                          = @{ EU=1; LV=20; SEQ=42 }

    # -- EU=1  LV=20  SEQ=44 : Account structure activation + Workflow queue -
    'Account structure activation'                    = @{ EU=1; LV=20; SEQ=44 }
    'Workflow work item queue'                        = @{ EU=1; LV=20; SEQ=44 }

    # -- EU=1  LV=20  SEQ=50 : Number sequences + Workflow queue assignee ----
    'Number sequence references'                      = @{ EU=1; LV=20; SEQ=50 }
    'Number sequence group'                           = @{ EU=1; LV=20; SEQ=50 }
    'Workflow work item queue assignee'               = @{ EU=1; LV=20; SEQ=50 }

    # -- EU=1  LV=20  SEQ=55 ------------------------------------------------
    'Workflow work item queue assignment'             = @{ EU=1; LV=20; SEQ=55 }
}

# =============================================================================
#  1.  Banner + pre-flight validation
# =============================================================================
Write-Banner
Write-Host ''
Write-Info "Environment  : $EnvironmentUrl"
Write-Info "Legal entity : $LegalEntityId"
Write-Info "Resources    : $ResourcesPath"
Write-Info "Output       : $OutputPath"

if (-not (Test-Path $ResourcesPath -PathType Container)) {
    throw "Resources path not found: '$ResourcesPath'"
}
if (-not (Test-Path $OutputPath -PathType Container)) {
    throw "Output path not found: '$OutputPath'"
}

# =============================================================================
#  2.  Discover packages
# =============================================================================
Write-Step 'Discovering packages'

$allFolders = @(Get-ChildItem -Path $ResourcesPath -Directory)
if ($allFolders.Count -eq 0) {
    throw "No subfolders found under '$ResourcesPath'.  Add at least one package folder."
}

Write-Info "Scanning $($allFolders.Count) folder(s)..."
$allPackages = for ($i = 0; $i -lt $allFolders.Count; $i++) {
    Get-PackageInfo -Folder $allFolders[$i] -Index ($i + 1)
}

# =============================================================================
#  3.  Package selection
# =============================================================================
$selectedPackages = [System.Collections.Generic.List[pscustomobject]]::new()

if ($PackageName) {
    # ── Single-package mode (non-interactive) ──────────────────────────────
    $match = $allPackages | Where-Object { $_.Name -eq $PackageName }
    if (-not $match) {
        $available = ($allPackages | ForEach-Object { "    '$($_.Name)'" }) -join [System.Environment]::NewLine
        throw "Package '$PackageName' not found under '$ResourcesPath'.`nAvailable packages:`n$available"
    }
    $selectedPackages.Add($match)
    Write-Info "Package : $PackageName"
} else {
    # ── Interactive numbered selection ─────────────────────────────────────
    Write-Step 'Select packages to import'

    $nameWidth = ($allPackages | ForEach-Object { $_.Name.Length } | Measure-Object -Maximum).Maximum
    $nameWidth = [Math]::Max(30, [Math]::Min($nameWidth, $Script:LineWidth - 26))

    $colFmt = '  {0,3}  {1}  {2,5}  {3,8}'
    Write-Host ''
    Write-Host ($colFmt -f '#', 'Package'.PadRight($nameWidth), 'Files', 'Entities') -ForegroundColor White
    Write-Host ($colFmt -f '---', '-' * $nameWidth, '-----', '--------') -ForegroundColor DarkGray

    foreach ($pkg in $allPackages) {
        $nameCol = if ($pkg.Name.Length -gt $nameWidth) {
                       $pkg.Name.Substring(0, $nameWidth - 3) + '...'
                   } else { $pkg.Name.PadRight($nameWidth) }
        $filesCol    = $pkg.XlsxCount.ToString().PadLeft(5)
        $entitiesCol = if ($pkg.HasManifest) { $pkg.EntityCount.ToString().PadLeft(8) } else { '       !' }
        $color       = if ($pkg.IsValid) { 'White' } else { 'DarkYellow' }
        Write-Host ($colFmt -f $pkg.Index, $nameCol, $filesCol, $entitiesCol) -ForegroundColor $color
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
        exit 0
    }

    foreach ($idx in $selectedIndices) {
        $selectedPackages.Add(($allPackages | Where-Object { $_.Index -eq $idx }))
    }
}

# =============================================================================
#  4.  Confirmation
# =============================================================================
Write-Host ''
Write-Rule 'Ready to import'
Write-Host ''
Write-Info "Environment  : $EnvironmentUrl"
Write-Info "Legal entity : $LegalEntityId"
Write-Info "Packages     : $($selectedPackages.Count)"

foreach ($pkg in $selectedPackages) {
    $meta = "$($pkg.XlsxCount) file(s), $($pkg.EntityCount) entities"
    if (-not $pkg.IsValid) { $meta += "  [!] $($pkg.Warnings -join '; ')" }
    Write-Detail "[$($pkg.Index)] $($pkg.Name)  ($meta)"
}

Write-Host ''
$confirm = (Read-Host '  Proceed? [Y]es / [N]o  (default: Y)').Trim().ToUpper()
if ($confirm -eq 'N') {
    Write-Info 'Cancelled.'
    exit 0
}

# =============================================================================
#  5.  Authenticate  (once -- token reused across all packages)
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
#  6.  Process packages
# =============================================================================
$results     = [System.Collections.Generic.List[pscustomobject]]::new()
$scriptStart = Get-Date

foreach ($pkg in $selectedPackages) {
    $pkgName     = $pkg.Name
    $pkgSafeName = ($pkgName -replace '[^A-Za-z0-9]', '-') -replace '-{2,}', '-'
    $pkgStart    = Get-Date
    $pkgResult   = [pscustomobject]@{
        Package     = $pkgName
        Status      = 'Error'
        ExecutionId = '-'
        Elapsed     = '-'
    }

    $divider = '=' * $Script:LineWidth
    Write-Host ''
    Write-Host $divider -ForegroundColor DarkCyan
    Write-Host "  [$($results.Count + 1)/$($selectedPackages.Count)]  $pkgName" -ForegroundColor Cyan
    Write-Host $divider -ForegroundColor DarkCyan

    # Token expiry warnings
    if ((Get-Date) -ge $tokenExpiry) {
        Write-Warn 'Access token has expired.  API calls will likely fail with HTTP 401.  Re-run the script.'
    } elseif ((Get-Date).AddMinutes(5) -ge $tokenExpiry) {
        Write-Warn "Token expires at $($tokenExpiry.ToString('HH:mm:ss')) -- it may expire mid-import."
    }

    try {
        # ── a. Discover xlsx files ────────────────────────────────────────────
        Write-Step 'Discovering files'
        $xlsxFiles = @(Get-ChildItem -Path $pkg.Folder.FullName -Filter '*.xlsx' -File)
        if ($xlsxFiles.Count -eq 0) { throw "No .xlsx files found in '$($pkg.Folder.FullName)'." }
        $xlsxNames = $xlsxFiles.Name
        Write-Info "$($xlsxFiles.Count) .xlsx file(s) found."

        # ── b. Parse and filter Manifest.xml ──────────────────────────────────
        Write-Step 'Filtering manifest'
        $manifestPath = Join-Path $pkg.Folder.FullName 'Manifest.xml'
        if (-not (Test-Path $manifestPath)) { throw "Manifest.xml not found in '$($pkg.Folder.FullName)'." }

        $existingDoc = New-Object System.Xml.XmlDocument
        $existingDoc.Load($manifestPath)
        $nsMgr = New-Object System.Xml.XmlNamespaceManager($existingDoc.NameTable)
        $nsMgr.AddNamespace('dm', $Script:DmNs)

        $matchedNodes = [System.Collections.Generic.List[System.Xml.XmlNode]]::new()
        $matchedFiles = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($node in $existingDoc.SelectNodes('//dm:DataManagementPackageEntityData', $nsMgr)) {
            $ifpNode = $node.SelectSingleNode('dm:InputFilePath', $nsMgr)
            if ($ifpNode -and ($xlsxNames -contains $ifpNode.InnerText)) {
                $matchedNodes.Add($node)
                [void]$matchedFiles.Add($ifpNode.InnerText)
            }
        }

        $unmanifestedFiles = @($xlsxNames | Where-Object { -not $matchedFiles.Contains($_) })
        if ($unmanifestedFiles.Count -gt 0) {
            Write-Warn "No manifest entry for: $($unmanifestedFiles -join ', ') -- these files will not be imported."
        }
        if ($matchedNodes.Count -eq 0) { throw 'No manifest entries matched the .xlsx files on disk.' }
        Write-Info "$($matchedNodes.Count) entity definition(s) matched."

        # ── c. Build optimised Manifest.xml ───────────────────────────────────
        Write-Step 'Building manifest'
        $timestamp         = Get-Date -Format 'yyyyMMddHHmmss'
        $guidSuffix        = [System.Guid]::NewGuid().ToString('N').Substring(0, 8).ToUpper()
        $definitionGroupId = "$pkgSafeName-$guidSuffix"
        Write-Info "Definition group : $definitionGroupId"

        $newDoc = New-Object System.Xml.XmlDocument
        [void]$newDoc.AppendChild($newDoc.CreateXmlDeclaration('1.0', 'utf-16', $null))
        $root = $newDoc.CreateElement('DataManagementPackageManifest', $Script:DmNs)
        $root.SetAttribute('xmlns:i', 'http://www.w3.org/2001/XMLSchema-instance')
        [void]$newDoc.AppendChild($root)

        foreach ($pair in @(
            [pscustomobject]@{ Name = 'DefinitionGroupName'; Value = $definitionGroupId },
            [pscustomobject]@{ Name = 'Description';          Value = $pkgName }
        )) {
            $el = $newDoc.CreateElement($pair.Name, $Script:DmNs)
            $el.InnerText = $pair.Value
            [void]$root.AppendChild($el)
        }

        $entityListEl = $newDoc.CreateElement('PackageEntityList', $Script:DmNs)
        [void]$root.AppendChild($entityListEl)
        $newNsMgr = New-Object System.Xml.XmlNamespaceManager($newDoc.NameTable)
        $newNsMgr.AddNamespace('dm', $Script:DmNs)

        $orderedCount = 0
        foreach ($node in $matchedNodes) {
            $imported = $newDoc.ImportNode($node, $true)
            $nameNode = $imported.SelectSingleNode('dm:EntityName', $newNsMgr)
            if ($nameNode -and $entityOrdering.Contains($nameNode.InnerText)) {
                $order = $entityOrdering[$nameNode.InnerText]
                $imported.SelectSingleNode('dm:ExecutionUnit',        $newNsMgr).InnerText = [string]$order.EU
                $imported.SelectSingleNode('dm:LevelInExecutionUnit', $newNsMgr).InnerText = [string]$order.LV
                $imported.SelectSingleNode('dm:SequenceInLevel',      $newNsMgr).InnerText = [string]$order.SEQ
                $orderedCount++
            }
            [void]$entityListEl.AppendChild($imported)
        }
        if ($orderedCount -gt 0) {
            Write-Info "Execution ordering applied to $orderedCount / $($matchedNodes.Count) entities."
        }

        # ── d. Build .zip package ─────────────────────────────────────────────
        Write-Step 'Building package zip'
        $tempDir = Join-Path $OutputPath "DMF_Build_${timestamp}_${pkgSafeName}"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

        try {
            # Manifest.xml -- UTF-16 LE with BOM (required by D365 DMF)
            $writerSettings          = New-Object System.Xml.XmlWriterSettings
            $writerSettings.Encoding = [System.Text.Encoding]::Unicode
            $writerSettings.Indent   = $true
            $xmlWriter = [System.Xml.XmlWriter]::Create((Join-Path $tempDir 'Manifest.xml'), $writerSettings)
            try   { $newDoc.Save($xmlWriter) }
            finally { $xmlWriter.Dispose() }

            # PackageHeader.xml -- UTF-16 LE with BOM
            $pkgNameXml  = [System.Security.SecurityElement]::Escape($pkgName)
            $headerXml   = (@(
                '<?xml version="1.0" encoding="utf-16"?>',
                '<DataManagementPackageHeader xmlns:i="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/dynamics/2015/01/DataManagement">',
                "  <Description>$pkgNameXml</Description>",
                '  <ManifestType>Microsoft.Dynamics.AX.Framework.Tools.DataManagement.Serialization.DataManagementPackageManifest</ManifestType>',
                '  <PackageType>DefinitionGroup</PackageType>',
                '  <PackageVersion>2</PackageVersion>',
                '</DataManagementPackageHeader>'
            ) -join [System.Environment]::NewLine)
            [System.IO.File]::WriteAllText(
                (Join-Path $tempDir 'PackageHeader.xml'),
                $headerXml,
                [System.Text.Encoding]::Unicode
            )

            foreach ($xlsx in $xlsxFiles) {
                Copy-Item -Path $xlsx.FullName -Destination (Join-Path $tempDir $xlsx.Name)
            }

            $zipPath   = Join-Path $OutputPath "${pkgSafeName}_${timestamp}.zip"
            [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $zipPath)
            $zipSizeMB = [Math]::Round((Get-Item $zipPath).Length / 1MB, 2)
            Write-Info "Package zip : $zipPath  ($zipSizeMB MB)"
        }
        finally {
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }

        # ── e. Request Azure Blob upload URL ──────────────────────────────────
        Write-Step 'Requesting blob upload URL'
        $blobFileName = "${pkgSafeName}_${timestamp}.zip"
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

        # ── f. Upload to Azure Blob ───────────────────────────────────────────
        Write-Step "Uploading package ($zipSizeMB MB)"
        $null = Invoke-DmfRequest -Operation 'blob upload' -Params @{
            Method      = 'Put'
            Uri         = $blobUrl
            Headers     = @{ 'x-ms-blob-type' = 'BlockBlob' }
            ContentType = 'application/octet-stream'
            InFile      = $zipPath
        }
        Write-Info 'Upload complete.'

        # ── g. Submit import job ──────────────────────────────────────────────
        Write-Step "Submitting import job  (legal entity: $LegalEntityId)"
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
                overwrite         = $true
                legalEntityId     = $LegalEntityId
            } | ConvertTo-Json)
        }

        $executionId = $importResp.value
        if ([string]::IsNullOrWhiteSpace($executionId)) {
            throw "ImportFromPackage returned an empty execution ID.  Response: $($importResp | ConvertTo-Json -Depth 5)"
        }
        Write-Info "Execution ID : $executionId"

        # ── h. Poll for completion (live progress bar) ────────────────────────
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
                    Write-Info "  [$( Get-Date -Format 'HH:mm:ss')]  Status: $currentStatus"
                    $lastStatus = $currentStatus
                }

                if ($terminalStatuses -contains $currentStatus) { break }
            }
        }
        finally {
            Write-Progress -Id 1 -Activity "Importing: $pkgName" -Completed
        }

        # ── i. Package outcome ────────────────────────────────────────────────
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
                Write-Warn 'Import partially succeeded - some records may have been skipped or errored.'
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
}

# =============================================================================
#  7.  Summary
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

if ($failCount -gt 0) { exit 1 }
