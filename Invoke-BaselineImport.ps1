#Requires -Version 5.1
<#
.SYNOPSIS
    Discovers D365 F&O data packages under the resources folder and imports them
    via the Data Management Framework API.

.DESCRIPTION
    1. Scans all subfolders under ResourcesPath, treating each as an importable package.
    2. Presents an interactive selection menu (skipped when -PackageName is supplied).
    3. Prompts for confirmation before authenticating (suppressed with -Force).
    4. Authenticates once via Microsoft Entra device code flow; the token is reused
       across all selected packages.
    5. For each confirmed package:
       a. Discovers .xlsx files on disk and cross-references them against Manifest.xml.
       b. Resolves entity execution ordering from a per-package ordering.json file (when
          present), merging it with the built-in default table; package-level entries
          take precedence for any entity listed in both.
       c. Rebuilds the manifest with the resolved ordering so independent entity chains
          can run in parallel within D365 DMF.
       d. Builds a well-formed .zip package and uploads it to Azure Blob Storage.
       e. Submits the import via ImportFromPackage and polls until a terminal status
          is reached, showing a live progress bar throughout.
       f. Deletes the generated .zip after upload (unless -KeepZip is specified).
    6. Prints a colour-coded summary table with per-package elapsed time and
       execution IDs for follow-up in D365 Job history.
    7. Writes a timestamped transcript to -LogPath (auto-generated when omitted;
       pass an empty string to suppress logging entirely).

    -WhatIf validates all selected packages and shows what would be imported without
    making any API calls to D365.

    All REST calls include automatic retry with exponential back-off (HTTP 5xx /
    429 / transient network errors).  The token expiry window is tracked and
    surfaced as a warning when approaching expiry.

    Library files (in ./lib/)
    ─────────────────────────
    DmfOutput.ps1   -- Write-* helpers, Format-Elapsed, Stop-RunTranscript
    DmfRequest.ps1  -- Invoke-DmfRequest (REST client with retry)
    DmfPackage.ps1  -- Get-PackageInfo, Resolve-EntityOrdering, $entityOrdering

    Per-package execution ordering
    ──────────────────────────────
    Place an ordering.json file inside any package folder to customise entity
    execution order for that package.  Entries in the file are merged with the
    built-in defaults; package-level values take precedence for any named entity.

    ordering.json format:
    {
      "Entity Name": { "EU": 1, "LV": 10, "SEQ": 10 },
      ...
    }

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
    Example: '010 - System Setup'

.PARAMETER ResourcesPath
    Root folder that contains package subfolders.
    Defaults to the repo-relative 'resources' directory next to this script.

.PARAMETER OutputPath
    Directory for generated .zip files.  Defaults to $env:TEMP.

.PARAMETER LogPath
    Path for the run transcript log.  When omitted, a log is auto-generated in
    OutputPath as DMFImport_<timestamp>.log.  Pass an empty string ('') to
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

.PARAMETER KeepZip
    Retain the generated .zip file after a successful upload.  By default the
    zip is deleted once the blob upload completes successfully.

.PARAMETER PassThru
    Return the per-package result objects to the pipeline after completion.
    Each object has properties: Package, Status, ExecutionId, Elapsed.

.EXAMPLE
    # Interactive: discover all packages and prompt for selection
    .\Invoke-BaselineImport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'DAT'

.EXAMPLE
    # Non-interactive: import one package, skip confirmation, write log
    .\Invoke-BaselineImport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'DAT' `
        -PackageName    '010 - System Setup' `
        -Force `
        -LogPath        'C:\Logs\import.log'

.EXAMPLE
    # Dry-run: validate all packages without making any API calls
    .\Invoke-BaselineImport.ps1 `
        -EnvironmentUrl 'https://contoso.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'DAT' `
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

    [string]$PackageName,

    [ValidateNotNullOrEmpty()]
    [string]$ResourcesPath = (Join-Path $PSScriptRoot 'resources'),

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = $env:TEMP,

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
    [switch]$KeepZip,
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
. (Join-Path $libPath 'DmfPackage.ps1')

# =============================================================================
#  Pre-flight path validation  (before transcript so errors surface cleanly)
# =============================================================================
if (-not (Test-Path $ResourcesPath -PathType Container)) {
    throw "Resources path not found: '$ResourcesPath'"
}
if (-not (Test-Path $OutputPath -PathType Container)) {
    throw "Output path not found: '$OutputPath'"
}

# =============================================================================
#  Script-level constants  (consumed by lib functions via $Script: scope)
# =============================================================================
$Script:Version          = '3.0'
$Script:DmNs             = 'http://schemas.microsoft.com/dynamics/2015/01/DataManagement'
$Script:LineWidth        = 80
$Script:MaxRetries       = $MaxRetries
$Script:TranscriptActive = $false

try {
    $w = $Host.UI.RawUI.BufferSize.Width
    $Script:LineWidth = [Math]::Min(120, [Math]::Max(80, $w))
} catch { <# non-interactive host -- keep default #> }

# =============================================================================
#  1.  Transcript startup  (OutputPath validated above)
# =============================================================================
if (-not $PSBoundParameters.ContainsKey('LogPath')) {
    $ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
    $LogPath = Join-Path $OutputPath "DMFImport_${ts}.log"
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
Write-Banner -DryRun:$WhatIf
Write-Host ''
Write-Info "Environment  : $EnvironmentUrl"
Write-Info "Tenant       : $TenantId"
Write-Info "Legal entity : $LegalEntityId"
Write-Info "Resources    : $ResourcesPath"
Write-Info "Output       : $OutputPath"
if ($Script:TranscriptActive) { Write-Info "Log          : $LogPath" }
if ($WhatIf)      { Write-Warn 'WhatIf active -- no API calls will be made.' }
if ($Force)       { Write-Info 'Force        : confirmation prompt suppressed' }
if ($NoOverwrite) { Write-Info 'NoOverwrite  : existing records will not be overwritten' }
if ($KeepZip)     { Write-Info 'KeepZip      : generated zip files will be retained' }

# =============================================================================
#  3.  Discover packages
# =============================================================================
Write-Step 'Discovering packages'

$allFolders = @(Get-ChildItem -Path $ResourcesPath -Directory)
if ($allFolders.Count -eq 0) {
    throw "No subfolders found under '$ResourcesPath'.  Add at least one package folder."
}

Write-Info "Scanning $($allFolders.Count) folder(s)..."
$allPackages = @(for ($i = 0; $i -lt $allFolders.Count; $i++) {
    Get-PackageInfo -Folder $allFolders[$i] -Index ($i + 1)
})

# =============================================================================
#  4.  Package selection
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
    $nameWidth = [int][Math]::Max(30, [Math]::Min($nameWidth, $Script:LineWidth - 36))

    $colFmt = '  {0,3}  {1}  {2,5}  {3,8}  {4,8}'
    Write-Host ''
    Write-Host ($colFmt -f '#', 'Package'.PadRight($nameWidth), 'Files', 'Size(MB)', 'Entities') -ForegroundColor White
    Write-Host ($colFmt -f '---', ('-' * $nameWidth), '-----', '--------', '--------') -ForegroundColor DarkGray

    foreach ($pkg in $allPackages) {
        $nameCol     = if ($pkg.Name.Length -gt $nameWidth) {
                           $pkg.Name.Substring(0, $nameWidth - 3) + '...'
                       } else { $pkg.Name.PadRight($nameWidth) }
        $filesCol    = $pkg.XlsxCount.ToString().PadLeft(5)
        $sizeCol     = if ($pkg.XlsxCount -gt 0) { $pkg.XlsxSizeMB.ToString('F2').PadLeft(8) } else { '       -' }
        $entitiesCol = if ($pkg.HasManifest) { $pkg.EntityCount.ToString().PadLeft(8) } else { '       !' }
        $color       = if ($pkg.IsValid) { 'White' } else { 'DarkYellow' }
        Write-Host ($colFmt -f $pkg.Index, $nameCol, $filesCol, $sizeCol, $entitiesCol) -ForegroundColor $color
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
#  5.  WhatIf exit  (validate and display -- no API calls made)
# =============================================================================
if ($WhatIf) {
    Write-Host ''
    Write-Rule 'WhatIf -- validation complete, no import will occur'
    Write-Host ''
    Write-Info "Environment  : $EnvironmentUrl"
    Write-Info "Legal entity : $LegalEntityId"
    Write-Info "Packages     : $($selectedPackages.Count)"
    foreach ($pkg in $selectedPackages) {
        $meta = "$($pkg.XlsxCount) file(s), $($pkg.XlsxSizeMB) MB, $($pkg.EntityCount) entities"
        if ($pkg.HasOrdering) { $meta += ', custom ordering.json' }
        if (-not $pkg.IsValid) { $meta += "  [!] $($pkg.Warnings -join '; ')" }
        Write-Detail "[$($pkg.Index)] $($pkg.Name)  ($meta)"
    }
    Write-Host ''
    Write-Info 'No API calls made.  Remove -WhatIf to perform the import.'
    if ($PassThru) {
        $selectedPackages | ForEach-Object {
            [pscustomobject]@{ Package = $_.Name; Status = 'WhatIf'; ExecutionId = '-'; Elapsed = '-' }
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
    $meta = "$($pkg.XlsxCount) file(s), $($pkg.XlsxSizeMB) MB, $($pkg.EntityCount) entities"
    if ($pkg.HasOrdering) { $meta += ', custom ordering.json' }
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
    $pkgName     = $pkg.Name
    $pkgSafeName = ($pkgName -replace '[^A-Za-z0-9]', '-') -replace '-{2,}', '-'
    $pkgStart    = Get-Date
    $pkgResult   = [pscustomobject]@{
        Package     = $pkgName
        Status      = 'Error'
        ExecutionId = '-'
        Elapsed     = '-'
    }
    $zipPath = $null

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

        # Resolve ordering: built-in defaults merged with any per-package ordering.json
        $effectiveOrdering = Resolve-EntityOrdering -PackageFolderPath $pkg.Folder.FullName

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
            if ($nameNode -and $effectiveOrdering.Contains($nameNode.InnerText)) {
                $order = $effectiveOrdering[$nameNode.InnerText]
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
            $pkgNameXml = [System.Security.SecurityElement]::Escape($pkgName)
            $headerXml  = (@(
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

        # -- Clean up zip once upload succeeds (unless -KeepZip) --------------
        if (-not $KeepZip) {
            Remove-Item -Path $zipPath -Force -ErrorAction SilentlyContinue
            Write-Detail 'Zip deleted.'
            $zipPath = $null
        } else {
            Write-Detail "Zip retained : $zipPath"
        }

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
                overwrite         = (-not $NoOverwrite.IsPresent)
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
                    Write-Info "  [$(Get-Date -Format 'HH:mm:ss')]  Status: $currentStatus"
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
    finally {
        # Ensure zip is removed on the error path when cleanup wasn't reached above
        if ($zipPath -and (Test-Path $zipPath) -and -not $KeepZip) {
            Remove-Item -Path $zipPath -Force -ErrorAction SilentlyContinue
        }
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
