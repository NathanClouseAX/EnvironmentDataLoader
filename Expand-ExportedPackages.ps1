#Requires -Version 5.1
<#
.SYNOPSIS
    Extracts downloaded D365 F&O DMF package .zip files into individual
    subfolders, ready for review and re-import via Invoke-BaselineImport.ps1.

.DESCRIPTION
    This script is the middle step of the export → review → re-import pipeline:

        Step 1 -- Export + Download
            .\Invoke-TemplateExport.ps1 ... -DownloadPath C:\DMF\Downloads

        Step 2 -- Extract for review  (this script)
            .\Expand-ExportedPackages.ps1 -SourcePath C:\DMF\Downloads `
                                          -DestinationPath C:\DMF\Packages

        Step 3 -- Review / edit the .xlsx files in C:\DMF\Packages\*

        Step 4 -- Re-package and import into target legal entity
            .\Invoke-BaselineImport.ps1 ... -ResourcesPath C:\DMF\Packages `
                                            -LegalEntityId TARGET

    For each selected .zip file the script will:
        1. Read the embedded Manifest.xml to determine the definition group name
           and entity count.
        2. Create a subfolder under DestinationPath named after the zip file
           (without the .zip extension).
        3. Extract all zip contents into that subfolder.
        4. Validate the result -- checks for Manifest.xml and at least one
           .xlsx file.

    If the destination subfolder already exists, the package is skipped unless
    -Force is specified (which deletes and re-extracts the folder).

    No D365 API calls are made; this script is entirely local.

.PARAMETER SourcePath
    Directory containing the downloaded DMF package .zip files.

.PARAMETER DestinationPath
    Root directory under which per-package subfolders will be created.
    Defaults to SourcePath when not specified.

.PARAMETER PackageName
    When supplied, extracts this one .zip file without prompting.
    The value may include or omit the .zip extension.

.PARAMETER Force
    Delete and re-extract the destination folder if it already exists.
    Without this switch, packages whose destination folder exists are skipped.

.PARAMETER WhatIf
    Show what would be extracted without creating any folders or files.

.PARAMETER PassThru
    Return per-package result objects to the pipeline.
    Each object has properties: Package, DefinitionGroupId, ExtractedTo,
    XlsxCount, EntityCount, Status.

.EXAMPLE
    # Interactive: list all .zip files and choose which to extract
    .\Expand-ExportedPackages.ps1 `
        -SourcePath      'C:\DMF\Downloads' `
        -DestinationPath 'C:\DMF\Packages'

.EXAMPLE
    # Extract everything, overwrite any existing folders, pipe results forward
    .\Expand-ExportedPackages.ps1 `
        -SourcePath      'C:\DMF\Downloads' `
        -DestinationPath 'C:\DMF\Packages' `
        -Force `
        -PassThru |
    ForEach-Object { Write-Host "$($_.Package) -> $($_.ExtractedTo)" }

.EXAMPLE
    # WhatIf: show what would be extracted without touching the filesystem
    .\Expand-ExportedPackages.ps1 `
        -SourcePath 'C:\DMF\Downloads' `
        -WhatIf
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$SourcePath,

    [string]$DestinationPath,

    [string]$PackageName,

    [switch]$Force,
    [switch]$WhatIf,
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
. (Join-Path $libPath 'DmfZip.ps1')

# =============================================================================
#  Pre-flight path validation
# =============================================================================
if (-not (Test-Path $SourcePath -PathType Container)) {
    throw "Source path not found: '$SourcePath'"
}

if (-not $PSBoundParameters.ContainsKey('DestinationPath') -or $DestinationPath -eq '') {
    $DestinationPath = $SourcePath
}

if (-not (Test-Path $DestinationPath -PathType Container)) {
    throw "Destination path not found: '$DestinationPath'"
}

# =============================================================================
#  Script-level constants
# =============================================================================
$Script:Version          = '1.0'
$Script:DmNs             = 'http://schemas.microsoft.com/dynamics/2015/01/DataManagement'
$Script:LineWidth        = 80
$Script:TranscriptActive = $false   # no transcript for this local-only script

try {
    $w = $Host.UI.RawUI.BufferSize.Width
    $Script:LineWidth = [Math]::Min(120, [Math]::Max(80, $w))
} catch { <# non-interactive host -- keep default #> }

# =============================================================================
#  1.  Banner + session info
# =============================================================================
Write-Banner -DryRun:$WhatIf -Title 'D365 F&O Package Extractor'
Write-Host ''
Write-Info "Source       : $SourcePath"
Write-Info "Destination  : $DestinationPath"
if ($WhatIf) { Write-Warn 'WhatIf active -- no files will be written.' }
if ($Force)  { Write-Info 'Force        : existing destination folders will be overwritten' }

# =============================================================================
#  2.  Discover packages
# =============================================================================
Write-Step 'Discovering packages'

$allZips = @(Get-ChildItem -Path $SourcePath -Filter '*.zip' -File)
if ($allZips.Count -eq 0) {
    throw "No .zip files found under '$SourcePath'."
}

Write-Info "Scanning $($allZips.Count) zip file(s)..."
$allPackages = @(for ($i = 0; $i -lt $allZips.Count; $i++) {
    Get-ZipPackageInfo -File $allZips[$i] -Index ($i + 1)
})

# =============================================================================
#  3.  Package selection
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
        throw "Package '$PackageName' not found under '$SourcePath'.`nAvailable packages:`n$available"
    }
    $selectedPackages.Add($match)
    Write-Info "Package : $($match.Name)"
} else {
    # ── Interactive numbered selection ─────────────────────────────────────
    Write-Step 'Select packages to extract'

    $nameWidth = ($allPackages | ForEach-Object { $_.Name.Length } | Measure-Object -Maximum).Maximum
    $nameWidth = [int][Math]::Max(30, [Math]::Min($nameWidth, $Script:LineWidth - 38))

    $colFmt = '  {0,3}  {1}  {2,8}  {3,8}  {4}'
    Write-Host ''
    Write-Host ($colFmt -f '#', 'Package'.PadRight($nameWidth), 'Size(MB)', 'Entities', 'Definition Group') -ForegroundColor White
    Write-Host ($colFmt -f '---', ('-' * $nameWidth), '--------', '--------', '----------------') -ForegroundColor DarkGray

    foreach ($pkg in $allPackages) {
        $nameCol = if ($pkg.Name.Length -gt $nameWidth) {
                       $pkg.Name.Substring(0, $nameWidth - 3) + '...'
                   } else { $pkg.Name.PadRight($nameWidth) }
        $sizeCol = $pkg.SizeMB.ToString('F2').PadLeft(8)
        $entCol  = if ($pkg.HasManifest) { $pkg.EntityCount.ToString().PadLeft(8) } else { '       !' }
        $dgCol   = if ($pkg.DefinitionGroupId) { $pkg.DefinitionGroupId } else { '(unknown)' }

        # Warn when the destination folder already exists
        $destFolder = Join-Path $DestinationPath ($pkg.Name -replace '\.zip$', '')
        $exists     = Test-Path $destFolder -PathType Container
        $color      = if (-not $pkg.IsValid) { 'DarkYellow' } elseif ($exists) { 'DarkCyan' } else { 'White' }
        $suffix     = if ($exists) { '  [exists]' } else { '' }

        Write-Host ($colFmt -f $pkg.Index, $nameCol, $sizeCol, $entCol, "$dgCol$suffix") -ForegroundColor $color
    }

    $invalidPkgs = @($allPackages | Where-Object { -not $_.IsValid })
    if ($invalidPkgs.Count -gt 0) {
        Write-Host ''
        foreach ($p in $invalidPkgs) {
            Write-Warn "  [!] '$($p.Name)': $($p.Warnings -join '; ')"
        }
    }

    $existingCount = @($allPackages | Where-Object {
        Test-Path (Join-Path $DestinationPath ($_.Name -replace '\.zip$', '')) -PathType Container
    }).Count
    if ($existingCount -gt 0 -and -not $Force) {
        Write-Host ''
        Write-Info "$existingCount package(s) shown in cyan already have a destination folder.  Use -Force to overwrite."
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
#  4.  WhatIf exit
# =============================================================================
if ($WhatIf) {
    Write-Host ''
    Write-Rule 'WhatIf -- no files will be written'
    Write-Host ''
    Write-Info "Source       : $SourcePath"
    Write-Info "Destination  : $DestinationPath"
    Write-Info "Packages     : $($selectedPackages.Count)"

    foreach ($pkg in $selectedPackages) {
        $destFolder = Join-Path $DestinationPath ($pkg.Name -replace '\.zip$', '')
        $exists     = Test-Path $destFolder -PathType Container
        $meta       = "$($pkg.SizeMB) MB, $($pkg.EntityCount) entities"
        if ($pkg.DefinitionGroupId) { $meta += ", def group: $($pkg.DefinitionGroupId)" }
        $action = if ($exists -and $Force) { 'overwrite' } elseif ($exists) { 'SKIP (exists, use -Force)' } else { 'extract' }
        Write-Detail "[$($pkg.Index)] $($pkg.Name)  ($meta)"
        Write-Detail "      -> $destFolder  [$action]"
    }

    Write-Host ''
    Write-Info 'No changes made.  Remove -WhatIf to perform the extraction.'
    if ($PassThru) {
        $selectedPackages | ForEach-Object {
            $destFolder = Join-Path $DestinationPath ($_.Name -replace '\.zip$', '')
            [pscustomobject]@{
                Package           = $_.Name
                DefinitionGroupId = $_.DefinitionGroupId
                ExtractedTo       = $destFolder
                XlsxCount         = 0
                EntityCount       = $_.EntityCount
                Status            = 'WhatIf'
            }
        }
    }
    exit 0
}

# =============================================================================
#  5.  Extract packages
# =============================================================================
$results    = [System.Collections.Generic.List[pscustomobject]]::new()
$totalStart = Get-Date

foreach ($pkg in $selectedPackages) {
    $pkgStart   = Get-Date
    $destFolder = Join-Path $DestinationPath ($pkg.Name -replace '\.zip$', '')
    $pkgResult  = [pscustomobject]@{
        Package           = $pkg.Name
        DefinitionGroupId = $pkg.DefinitionGroupId
        ExtractedTo       = $destFolder
        XlsxCount         = 0
        EntityCount       = $pkg.EntityCount
        Status            = 'Error'
    }

    $divider = '=' * $Script:LineWidth
    Write-Host ''
    Write-Host $divider -ForegroundColor DarkCyan
    Write-Host "  [$($results.Count + 1)/$($selectedPackages.Count)]  $($pkg.Name)" -ForegroundColor Cyan
    if ($pkg.DefinitionGroupId) {
        Write-Host "  Definition group: $($pkg.DefinitionGroupId)" -ForegroundColor DarkGray
    }
    Write-Host $divider -ForegroundColor DarkCyan

    try {
        if (-not $pkg.IsValid) {
            throw "Package validation failed: $($pkg.Warnings -join '; ')"
        }

        # ── Check destination folder ──────────────────────────────────────────
        if (Test-Path $destFolder -PathType Container) {
            if (-not $Force) {
                Write-Warn "Destination folder already exists -- skipping.  Use -Force to overwrite."
                Write-Info "Folder : $destFolder"
                $pkgResult.Status = 'Skipped'
                $results.Add($pkgResult)
                if ($PassThru) { Write-Output $pkgResult }
                continue
            }
            Write-Info "Removing existing folder (Force)..."
            Remove-Item -Path $destFolder -Recurse -Force
        }

        # ── Create destination folder and extract ─────────────────────────────
        Write-Step 'Extracting'
        New-Item -ItemType Directory -Path $destFolder -Force | Out-Null

        [System.IO.Compression.ZipFile]::ExtractToDirectory($pkg.File.FullName, $destFolder)

        # ── Validate extraction ───────────────────────────────────────────────
        $extractedManifest = Join-Path $destFolder 'Manifest.xml'
        if (-not (Test-Path $extractedManifest)) {
            throw 'Manifest.xml was not found in the extracted folder.'
        }

        $xlsxFiles = @(Get-ChildItem -Path $destFolder -Filter '*.xlsx' -File)
        $pkgResult.XlsxCount = $xlsxFiles.Count
        if ($xlsxFiles.Count -eq 0) {
            Write-Warn 'No .xlsx files found in extracted folder -- package may be empty or malformed.'
        }

        $pkgResult.Status = 'Extracted'
        $elapsed = Format-Elapsed ((Get-Date) - $pkgStart)

        Write-OK "Extracted to : $destFolder  ($($xlsxFiles.Count) xlsx file(s), $($pkg.EntityCount) entities, $elapsed)"
        Write-Info "Next step    : review files, then run Invoke-BaselineImport.ps1 -ResourcesPath '$DestinationPath'"
    }
    catch {
        $pkgResult.Status = 'Error'
        Write-Fail "Failed to extract '$($pkg.Name)':"
        Write-Fail "  $_"
        Write-Verbose $_.ScriptStackTrace

        # Clean up a partial extraction folder on error
        if ((Test-Path $destFolder) -and $pkgResult.Status -eq 'Error') {
            try { Remove-Item -Path $destFolder -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        }
    }

    $results.Add($pkgResult)
    if ($PassThru) { Write-Output $pkgResult }
}

# =============================================================================
#  6.  Summary
# =============================================================================
$totalElapsed = Format-Elapsed ((Get-Date) - $totalStart)
$divider      = '=' * $Script:LineWidth

Write-Host ''
Write-Host $divider -ForegroundColor DarkCyan
Write-Host "  Summary  --  $($results.Count) package(s)  |  total time: $totalElapsed" -ForegroundColor Cyan
Write-Host $divider -ForegroundColor DarkCyan
Write-Host ''

$nameWidth2  = ($results | ForEach-Object { $_.Package.Length } | Measure-Object -Maximum).Maximum
$nameWidth2  = [Math]::Max(20, [Math]::Min($nameWidth2, $Script:LineWidth - 40))
$statusWidth = 10

$summaryHdr = '  {0}  {1}  {2,5}  {3}' -f 'Status'.PadRight($statusWidth), 'Package'.PadRight($nameWidth2), 'Files', 'Extracted To'
$summarySep = '  {0}  {1}  {2,5}  {3}' -f ('-' * $statusWidth), ('-' * $nameWidth2), '-----', '------------'

Write-Host $summaryHdr -ForegroundColor White
Write-Host $summarySep -ForegroundColor DarkGray

$extractedCount = 0
$skippedCount   = 0
$failCount      = 0

foreach ($r in $results) {
    $color = switch ($r.Status) {
        'Extracted' { $extractedCount++; 'Green'      }
        'Skipped'   { $skippedCount++;  'DarkCyan'   }
        default     { $failCount++;     'Red'         }
    }

    $nameTrunc = if ($r.Package.Length -gt $nameWidth2) {
        $r.Package.Substring(0, $nameWidth2 - 3) + '...'
    } else { $r.Package.PadRight($nameWidth2) }

    $row = '  {0}  {1}  {2,5}  {3}' -f $r.Status.PadRight($statusWidth), $nameTrunc, $r.XlsxCount, $r.ExtractedTo
    Write-Host $row -ForegroundColor $color
}

Write-Host ''
Write-Host $divider -ForegroundColor DarkCyan

if ($failCount -eq 0) {
    $msg   = if ($extractedCount -gt 0) {
        "  $extractedCount extracted$(if ($skippedCount) { ", $skippedCount skipped" }).  " +
        "Run Invoke-BaselineImport.ps1 -ResourcesPath '$DestinationPath' to import."
    } else {
        "  All $skippedCount package(s) skipped (folders already exist).  Use -Force to overwrite."
    }
    Write-Host $msg -ForegroundColor Green
} else {
    Write-Host "  $extractedCount extracted  |  $skippedCount skipped  |  $failCount failed." -ForegroundColor Red
}

Write-Host $divider -ForegroundColor DarkCyan
Write-Host ''

if ($failCount -gt 0) { exit 1 }
