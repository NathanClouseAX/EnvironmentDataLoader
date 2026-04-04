#Requires -Version 5.1
<#
.SYNOPSIS
    Discovers D365 F&O data packages under the resources folder and imports them
    via the Data Management API.

.DESCRIPTION
    1. Discovers all subfolders under ResourcesPath, each treated as a data package.
    2. Prompts the user to confirm which packages to import (skipped if -PackageName
       is specified).
    3. Authenticates once with Microsoft Entra using device code flow.
    4. For each confirmed package:
       a. Discovers .xlsx files and filters the package Manifest.xml to matched files.
       b. Rebuilds the manifest with optimized execution ordering.
       c. Zips the package and uploads it to Azure Blob Storage via GetAzureWriteUrl.
       d. Submits an import job via ImportFromPackage and polls until completion.
    5. Prints a per-package outcome summary.

.PARAMETER EnvironmentUrl
    Base URL of the D365 F&O environment.
    Example: https://myenv.operations.dynamics.com

.PARAMETER TenantId
    Microsoft Entra tenant ID (GUID) or domain name.
    Example: contoso.onmicrosoft.com

.PARAMETER LegalEntityId
    Legal entity to import into.
    Example: DAT

.PARAMETER PackageName
    Name of a single package folder to import without prompting. When omitted, all
    subfolders under ResourcesPath are discovered and the user is prompted for each.
    Example: 'Demo data 10.0.9 100 System and Shared'

.PARAMETER ResourcesPath
    Path to the folder containing package subfolders.
    Defaults to the repo-relative 'resources' folder.

.PARAMETER OutputPath
    Directory where generated .zip files will be written. Defaults to $env:TEMP.

.PARAMETER PollIntervalSeconds
    Seconds between status checks while waiting for import completion. Default: 30.

.PARAMETER TimeoutMinutes
    Maximum minutes to wait for a single package import before giving up. Default: 60.

.EXAMPLE
    # Import all packages, prompted per folder:
    .\Invoke-BaselineImport.ps1 `
        -EnvironmentUrl 'https://myenv.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'DAT'

.EXAMPLE
    # Import one specific package without prompting:
    .\Invoke-BaselineImport.ps1 `
        -EnvironmentUrl 'https://myenv.operations.dynamics.com' `
        -TenantId       'contoso.onmicrosoft.com' `
        -LegalEntityId  'DAT' `
        -PackageName    'Demo data 10.0.9 100 System and Shared'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
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
    [int]$TimeoutMinutes = 60
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression.FileSystem

# ─── Console helpers ──────────────────────────────────────────────────────────
function Write-Step { param([string]$Message) Write-Host "`n==> $Message" -ForegroundColor Cyan }
function Write-Info { param([string]$Message) Write-Host "    $Message" -ForegroundColor Gray }
function Write-OK   { param([string]$Message) Write-Host "`n[OK]   $Message" -ForegroundColor Green }
function Write-Warn { param([string]$Message) Write-Host "[WARN] $Message" -ForegroundColor Yellow }
function Write-Fail { param([string]$Message) Write-Host "[FAIL] $Message" -ForegroundColor Red }

# ─── Entity execution ordering ────────────────────────────────────────────────
# Entities sharing the same EU+LV+SEQ run in parallel.
# Within an EU, levels execute sequentially (lower first).
# Within a level, sequences execute sequentially (lower first).
#
#   EU=1  LV=10  Foundation + org/geographic data (must precede finance/workflow)
#   EU=1  LV=20  Finance (GL setup) and Workflow — merged into one level so their
#                independent chains run in parallel at each matching SEQ step.
#
# Entities not listed here fall back to the source Manifest.xml values.
#
$entityOrdering = [ordered]@{
    # ── EU=1  LV=10  SEQ=10 : Independent reference tables ────────────────────
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
    'Country/regions'                                 = @{ EU=1; LV=10; SEQ=10 }  # was SEQ=30
    'Team types'                                      = @{ EU=1; LV=10; SEQ=10 }  # was SEQ=200
    'System parameters'                               = @{ EU=1; LV=10; SEQ=10 }  # was SEQ=30

    # ── EU=1  LV=10  SEQ=20 : Depends on SEQ=10 ──────────────────────────────
    'Exchange rates'                                  = @{ EU=1; LV=10; SEQ=20 }
    'Rating level'                                    = @{ EU=1; LV=10; SEQ=20 }
    'Organization hierarchy purposes'                 = @{ EU=1; LV=10; SEQ=20 }
    'Address format lines'                            = @{ EU=1; LV=10; SEQ=20 }
    'Skills'                                          = @{ EU=1; LV=10; SEQ=20 }  # was SEQ=30
    'States'                                          = @{ EU=1; LV=10; SEQ=20 }  # was SEQ=40; Country/regions now at SEQ=10
    'Units'                                           = @{ EU=1; LV=10; SEQ=20 }  # was SEQ=40
    'Legal entities'                                  = @{ EU=1; LV=10; SEQ=20 }  # was SEQ=70; needs only Currencies+Country/regions

    # ── EU=1  LV=10  SEQ=30 : Depends on SEQ=20 ──────────────────────────────
    'Counties'                                        = @{ EU=1; LV=10; SEQ=30 }  # was SEQ=45
    'Unit translations'                               = @{ EU=1; LV=10; SEQ=30 }  # was SEQ=50
    'Unit conversions'                                = @{ EU=1; LV=10; SEQ=30 }  # was SEQ=50
    'Cities'                                          = @{ EU=1; LV=10; SEQ=30 }  # was SEQ=50; States now at SEQ=20
    'Operating unit'                                  = @{ EU=1; LV=10; SEQ=30 }  # was SEQ=80; Legal entities now at SEQ=20

    # ── EU=1  LV=10  SEQ=40 : Depends on SEQ=30 ──────────────────────────────
    'Districts V2'                                    = @{ EU=1; LV=10; SEQ=40 }  # was SEQ=55

    # ── EU=1  LV=10  SEQ=50 : Depends on SEQ=40 ──────────────────────────────
    'Postal codes V3'                                 = @{ EU=1; LV=10; SEQ=50 }  # was SEQ=60
    'Global address book V2'                          = @{ EU=1; LV=10; SEQ=50 }  # was SEQ=210; Operating unit now at SEQ=30

    # ── EU=1  LV=10  SEQ=60 : Depends on SEQ=50 ──────────────────────────────
    'Organization hierarchy V2 - published and draft' = @{ EU=1; LV=10; SEQ=60 }  # was SEQ=220
    'User information'                                = @{ EU=1; LV=10; SEQ=60 }  # was SEQ=230

    # ── EU=1  LV=10  SEQ=70 : Depends on SEQ=60 ──────────────────────────────
    'User to person relationship'                     = @{ EU=1; LV=10; SEQ=70 }  # was SEQ=240
    'Teams V2'                                        = @{ EU=1; LV=10; SEQ=70 }  # was SEQ=240

    # ── EU=1  LV=10  SEQ=80 : Depends on SEQ=70 ──────────────────────────────
    'Security user role association'                  = @{ EU=1; LV=10; SEQ=80 }  # was SEQ=250

    # ── EU=1  LV=10  SEQ=90 : Depends on SEQ=80 ──────────────────────────────
    'Party relationships'                             = @{ EU=1; LV=10; SEQ=90 }  # was SEQ=300

    # ── EU=1  LV=10  SEQ=100 : Depends on SEQ=90 ─────────────────────────────
    'Party contacts'                                  = @{ EU=1; LV=10; SEQ=100 }  # was SEQ=310

    # ── EU=1  LV=10  SEQ=110 : Depends on SEQ=100 ────────────────────────────
    'Party postal address V2'                         = @{ EU=1; LV=10; SEQ=110 }  # was SEQ=320

    # ── EU=1  LV=20  SEQ=10 : Finance foundation + Workflow foundation ────────
    # These two chains have no cross-dependencies and run in parallel per SEQ step.
    'Chart of accounts'                               = @{ EU=1; LV=20; SEQ=10 }
    'Fiscal calendar'                                 = @{ EU=1; LV=20; SEQ=10 }
    'Financial dimensions'                            = @{ EU=1; LV=20; SEQ=10 }
    'Main account categories'                         = @{ EU=1; LV=20; SEQ=10 }
    'Expression'                                      = @{ EU=1; LV=20; SEQ=10 }  # was LV=22 SEQ=10
    'System email template'                           = @{ EU=1; LV=20; SEQ=10 }  # was LV=22 SEQ=10

    # ── EU=1  LV=20  SEQ=20 ───────────────────────────────────────────────────
    'Main account'                                    = @{ EU=1; LV=20; SEQ=20 }
    'Fiscal calendar period'                          = @{ EU=1; LV=20; SEQ=20 }
    'Dimension attribute activation'                  = @{ EU=1; LV=20; SEQ=20 }
    'Financial dimension format'                      = @{ EU=1; LV=20; SEQ=20 }
    'Financial dimension translations'                = @{ EU=1; LV=20; SEQ=20 }
    'Workflow version'                                = @{ EU=1; LV=20; SEQ=20 }  # was LV=22 SEQ=20
    'Workflow parallel branch'                        = @{ EU=1; LV=20; SEQ=20 }  # was LV=22 SEQ=20
    'System email template message'                   = @{ EU=1; LV=20; SEQ=20 }  # was LV=22 SEQ=20

    # ── EU=1  LV=20  SEQ=30 ───────────────────────────────────────────────────
    'Advanced rule structures'                        = @{ EU=1; LV=20; SEQ=30 }
    'Consolidation groups and accounts'               = @{ EU=1; LV=20; SEQ=30 }
    'Financial dimension values'                      = @{ EU=1; LV=20; SEQ=30 }
    'Financial dimension value translations'          = @{ EU=1; LV=20; SEQ=30 }
    'Financial dimension sets'                        = @{ EU=1; LV=20; SEQ=30 }
    'Workflow version notes'                          = @{ EU=1; LV=20; SEQ=30 }  # was LV=22 SEQ=30
    'Workflow subworkflow'                            = @{ EU=1; LV=20; SEQ=30 }  # was LV=22 SEQ=30
    'Workflow system parameters'                      = @{ EU=1; LV=20; SEQ=30 }  # was LV=22 SEQ=30
    'Workflow element'                                = @{ EU=1; LV=20; SEQ=30 }  # was LV=22 SEQ=30

    # ── EU=1  LV=20  SEQ=32 ───────────────────────────────────────────────────
    'Advanced rule structure allowed values'          = @{ EU=1; LV=20; SEQ=32 }

    # ── EU=1  LV=20  SEQ=34 ───────────────────────────────────────────────────
    'Advanced rule structure activation'              = @{ EU=1; LV=20; SEQ=34 }

    # ── EU=1  LV=20  SEQ=36 : Account structures + Workflow element details ───
    'Account structures'                              = @{ EU=1; LV=20; SEQ=36 }
    'Workflow element action'                         = @{ EU=1; LV=20; SEQ=36 }  # was LV=22 SEQ=40
    'Workflow step'                                   = @{ EU=1; LV=20; SEQ=36 }  # was LV=22 SEQ=40
    'Workflow element notification'                   = @{ EU=1; LV=20; SEQ=36 }  # was LV=22 SEQ=40
    'Workflow element link'                           = @{ EU=1; LV=20; SEQ=36 }  # was LV=22 SEQ=40
    'Workflow element outcome message'                = @{ EU=1; LV=20; SEQ=36 }  # was LV=22 SEQ=40
    'Workflow version notification'                   = @{ EU=1; LV=20; SEQ=36 }  # was LV=22 SEQ=40

    # ── EU=1  LV=20  SEQ=38 ───────────────────────────────────────────────────
    'Account structure allowed values'                = @{ EU=1; LV=20; SEQ=38 }

    # ── EU=1  LV=20  SEQ=40 : Advanced rules + Workflow notifications ─────────
    'Advanced rules'                                  = @{ EU=1; LV=20; SEQ=40 }
    'Number sequence code'                            = @{ EU=1; LV=20; SEQ=40 }
    'Workflow version notification message'           = @{ EU=1; LV=20; SEQ=40 }  # was LV=22 SEQ=50
    'Workflow escalation path'                        = @{ EU=1; LV=20; SEQ=40 }  # was LV=22 SEQ=50
    'Workflow element notification message'           = @{ EU=1; LV=20; SEQ=40 }  # was LV=22 SEQ=50
    'Workflow step message'                           = @{ EU=1; LV=20; SEQ=40 }  # was LV=22 SEQ=50
    'Workflow line item'                              = @{ EU=1; LV=20; SEQ=40 }  # was LV=22 SEQ=50

    # ── EU=1  LV=20  SEQ=42 ───────────────────────────────────────────────────
    'Advanced rule criteria'                          = @{ EU=1; LV=20; SEQ=42 }

    # ── EU=1  LV=20  SEQ=44 : Account structure activation + Workflow queue ───
    'Account structure activation'                    = @{ EU=1; LV=20; SEQ=44 }
    'Workflow work item queue'                        = @{ EU=1; LV=20; SEQ=44 }  # was LV=22 SEQ=60

    # ── EU=1  LV=20  SEQ=50 : Number sequences + Workflow queue assignee ──────
    'Number sequence references'                      = @{ EU=1; LV=20; SEQ=50 }
    'Number sequence group'                           = @{ EU=1; LV=20; SEQ=50 }
    'Workflow work item queue assignee'               = @{ EU=1; LV=20; SEQ=50 }  # was LV=22 SEQ=65

    # ── EU=1  LV=20  SEQ=55 ───────────────────────────────────────────────────
    'Workflow work item queue assignment'             = @{ EU=1; LV=20; SEQ=55 }  # was LV=22 SEQ=70
}

$dmNs = 'http://schemas.microsoft.com/dynamics/2015/01/DataManagement'

# ─── 1. Discover package folders ──────────────────────────────────────────────
Write-Step "Discovering packages"
Write-Info "Resources: $ResourcesPath"

if (-not (Test-Path $ResourcesPath -PathType Container)) {
    throw "Resources path not found: '$ResourcesPath'"
}

$allPackages = @(Get-ChildItem -Path $ResourcesPath -Directory)
if ($allPackages.Count -eq 0) {
    throw "No package folders found under '$ResourcesPath'."
}

if ($PackageName) {
    $allPackages = @($allPackages | Where-Object { $_.Name -eq $PackageName })
    if ($allPackages.Count -eq 0) {
        throw "Package folder '$PackageName' not found under '$ResourcesPath'."
    }
}

# ─── 2. Select packages to import ─────────────────────────────────────────────
$selectedPackages = [System.Collections.Generic.List[System.IO.DirectoryInfo]]::new()

if ($PackageName) {
    $selectedPackages.Add($allPackages[0])
    Write-Info "Package: $PackageName"
} else {
    Write-Info "$($allPackages.Count) package(s) available:"
    foreach ($p in $allPackages) { Write-Info "  $($p.Name)" }

    $selectAll = $false
    :SelectionLoop foreach ($folder in $allPackages) {
        if ($selectAll) {
            $selectedPackages.Add($folder)
            continue
        }

        do {
            $answer = (Read-Host "`n  Import '$($folder.Name)'? [Y]es / [N]o / [A]ll / [Q]uit").Trim().ToUpper()
        } while ($answer -notin @('Y', 'N', 'A', 'Q', ''))

        switch ($answer) {
            'A' { $selectedPackages.Add($folder); $selectAll = $true }
            'Q' { break SelectionLoop }
            'N' { <# skip #> }
            default { $selectedPackages.Add($folder) }  # Y or Enter
        }
    }
}

if ($selectedPackages.Count -eq 0) {
    Write-Warn "No packages selected. Exiting."
    exit 0
}

Write-Info "$($selectedPackages.Count) package(s) queued for import."

# ─── 3. Authenticate (once for all packages) ──────────────────────────────────
Write-Step "Authenticating with Microsoft Entra (device code flow)"

$clientId = '1950a258-227b-4e31-a9cf-717495945fc2'
$baseUrl  = $EnvironmentUrl.TrimEnd('/')
$scope    = "$baseUrl/.default"
$authBase = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0"

$deviceCode = Invoke-RestMethod -Method Post `
    -Uri         "$authBase/devicecode" `
    -ContentType 'application/x-www-form-urlencoded' `
    -Body        "client_id=$clientId&scope=$([System.Uri]::EscapeDataString($scope))"

Write-Host "`n$($deviceCode.message)" -ForegroundColor Yellow

$pollUntil   = (Get-Date).AddSeconds($deviceCode.expires_in)
$accessToken = $null

while ((Get-Date) -lt $pollUntil) {
    Start-Sleep -Seconds $deviceCode.interval
    try {
        $tokenResp = Invoke-RestMethod -Method Post `
            -Uri         "$authBase/token" `
            -ContentType 'application/x-www-form-urlencoded' `
            -Body        "grant_type=urn:ietf:params:oauth:grant-type:device_code&client_id=$clientId&device_code=$($deviceCode.device_code)"
        $accessToken = $tokenResp.access_token
        Write-Info "Sign-in successful."
        break
    }
    catch {
        $errCode = $null
        try { $errCode = ($_.ErrorDetails.Message | ConvertFrom-Json).error } catch {}
        switch ($errCode) {
            'authorization_pending' { continue }
            'authorization_declined' { throw 'Sign-in was declined. Re-run and approve the request.' }
            'expired_token'          { throw 'The device code expired before sign-in completed. Re-run the script.' }
            default                  { throw }
        }
    }
}

if (-not $accessToken) {
    throw 'Authentication timed out before sign-in completed.'
}

$authHeaders = @{ Authorization = "Bearer $accessToken" }
$dmfBase     = "$baseUrl/data/DataManagementDefinitionGroups/Microsoft.Dynamics.DataEntities"

# ─── 4. Process each package ──────────────────────────────────────────────────
$results = [System.Collections.Generic.List[pscustomobject]]::new()

foreach ($pkg in $selectedPackages) {
    $pkgName     = $pkg.Name
    $pkgSafeName = ($pkgName -replace '[^A-Za-z0-9]', '-') -replace '-{2,}', '-'
    $pkgIdx      = $results.Count + 1
    $pkgResult   = [pscustomobject]@{ Package = $pkgName; Status = 'Error'; ExecutionId = '-' }

    Write-Host "`n$('─' * 72)" -ForegroundColor DarkGray
    Write-Step "Package $pkgIdx of $($selectedPackages.Count): $pkgName"

    try {
        # ── a. Discover xlsx files ───────────────────────────────────────────
        $xlsxFiles = @(Get-ChildItem -Path $pkg.FullName -Filter '*.xlsx' -File)
        if ($xlsxFiles.Count -eq 0) {
            throw "No .xlsx files found in '$($pkg.FullName)'."
        }
        $xlsxNames = $xlsxFiles.Name
        Write-Info "$($xlsxFiles.Count) file(s) found."

        # ── b. Parse and filter Manifest.xml ────────────────────────────────
        $manifestPath = Join-Path $pkg.FullName 'Manifest.xml'
        if (-not (Test-Path $manifestPath)) {
            throw "Manifest.xml not found in '$($pkg.FullName)'."
        }

        $existingDoc = New-Object System.Xml.XmlDocument
        $existingDoc.Load($manifestPath)
        $nsMgr = New-Object System.Xml.XmlNamespaceManager($existingDoc.NameTable)
        $nsMgr.AddNamespace('dm', $dmNs)

        $matchedNodes = [System.Collections.Generic.List[System.Xml.XmlNode]]::new()
        $matchedFiles = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($node in $existingDoc.SelectNodes('//dm:DataManagementPackageEntityData', $nsMgr)) {
            $ifpNode = $node.SelectSingleNode('dm:InputFilePath', $nsMgr)
            if ($ifpNode -and ($xlsxNames -contains $ifpNode.InnerText)) {
                $matchedNodes.Add($node)
                [void]$matchedFiles.Add($ifpNode.InnerText)
            }
        }

        $unmanifestedFiles = $xlsxNames | Where-Object { -not $matchedFiles.Contains($_) }
        if ($unmanifestedFiles) {
            Write-Warn "No manifest entry for: $($unmanifestedFiles -join ', '). These files will not be imported."
        }
        if ($matchedNodes.Count -eq 0) {
            throw "No manifest entries matched the xlsx files on disk."
        }
        Write-Info "$($matchedNodes.Count) entity definition(s) matched."

        # ── c. Build new Manifest.xml ────────────────────────────────────────
        $timestamp         = Get-Date -Format 'yyyyMMddHHmmss'
        $definitionGroupId = "$pkgSafeName-$([System.Guid]::NewGuid().ToString('N').Substring(0,8).ToUpper())"
        Write-Info "DefinitionGroupName: $definitionGroupId"

        $newDoc = New-Object System.Xml.XmlDocument
        [void]$newDoc.AppendChild($newDoc.CreateXmlDeclaration('1.0', 'utf-16', $null))
        $root = $newDoc.CreateElement('DataManagementPackageManifest', $dmNs)
        $root.SetAttribute('xmlns:i', 'http://www.w3.org/2001/XMLSchema-instance')
        [void]$newDoc.AppendChild($root)

        foreach ($pair in @(
            [pscustomobject]@{ Name = 'DefinitionGroupName'; Value = $definitionGroupId },
            [pscustomobject]@{ Name = 'Description';          Value = $pkgName }
        )) {
            $el = $newDoc.CreateElement($pair.Name, $dmNs)
            $el.InnerText = $pair.Value
            [void]$root.AppendChild($el)
        }

        $entityListEl = $newDoc.CreateElement('PackageEntityList', $dmNs)
        [void]$root.AppendChild($entityListEl)
        $newNsMgr = New-Object System.Xml.XmlNamespaceManager($newDoc.NameTable)
        $newNsMgr.AddNamespace('dm', $dmNs)

        foreach ($node in $matchedNodes) {
            $imported   = $newDoc.ImportNode($node, $true)
            $nameNode   = $imported.SelectSingleNode('dm:EntityName', $newNsMgr)
            if ($nameNode -and $entityOrdering.Contains($nameNode.InnerText)) {
                $order = $entityOrdering[$nameNode.InnerText]
                $imported.SelectSingleNode('dm:ExecutionUnit',        $newNsMgr).InnerText = [string]$order.EU
                $imported.SelectSingleNode('dm:LevelInExecutionUnit', $newNsMgr).InnerText = [string]$order.LV
                $imported.SelectSingleNode('dm:SequenceInLevel',      $newNsMgr).InnerText = [string]$order.SEQ
            }
            [void]$entityListEl.AppendChild($imported)
        }

        # ── d. Build zip ─────────────────────────────────────────────────────
        Write-Step "Building package zip"
        $tempDir = Join-Path $OutputPath "DMF_Build_$timestamp"
        New-Item -ItemType Directory -Path $tempDir | Out-Null

        try {
            $writerSettings          = New-Object System.Xml.XmlWriterSettings
            $writerSettings.Encoding = [System.Text.Encoding]::Unicode
            $writerSettings.Indent   = $true
            $xmlWriter = [System.Xml.XmlWriter]::Create((Join-Path $tempDir 'Manifest.xml'), $writerSettings)
            $newDoc.Save($xmlWriter)
            $xmlWriter.Close()

            $headerXml = @"
<?xml version="1.0" encoding="utf-16"?>
<DataManagementPackageHeader xmlns:i="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/dynamics/2015/01/DataManagement">
  <Description>$pkgName</Description>
  <ManifestType>Microsoft.Dynamics.AX.Framework.Tools.DataManagement.Serialization.DataManagementPackageManifest</ManifestType>
  <PackageType>DefinitionGroup</PackageType>
  <PackageVersion>2</PackageVersion>
</DataManagementPackageHeader>
"@
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
            $zipSizeMB = [math]::Round((Get-Item $zipPath).Length / 1MB, 2)
            Write-Info "Output: $zipPath ($zipSizeMB MB)"
        }
        finally {
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }

        # ── e. Get blob write URL ─────────────────────────────────────────────
        Write-Step "Requesting Azure Blob write URL"
        $blobFileName = "${pkgSafeName}_${timestamp}.zip"

        $writeUrlResp = Invoke-RestMethod -Method Post `
            -Uri         "$dmfBase.GetAzureWriteUrl" `
            -Headers     $authHeaders `
            -ContentType 'application/json' `
            -Body        (@{ uniqueFileName = $blobFileName } | ConvertTo-Json)

        # This endpoint returns Edm.String — the payload is a JSON-encoded string inside 'value'.
        $blobResult = $writeUrlResp.value | ConvertFrom-Json
        $blobUrl    = $blobResult.BlobUrl
        $blobId     = $blobResult.BlobId
        Write-Info "Blob ID: $blobId"

        if ([string]::IsNullOrWhiteSpace($blobUrl)) {
            throw "GetAzureWriteUrl did not return a BlobUrl. Response: $($writeUrlResp | ConvertTo-Json -Depth 5)"
        }

        # ── f. Upload zip ─────────────────────────────────────────────────────
        Write-Step "Uploading package ($zipSizeMB MB)"
        Invoke-RestMethod -Method Put `
            -Uri         $blobUrl `
            -Headers     @{ 'x-ms-blob-type' = 'BlockBlob' } `
            -ContentType 'application/octet-stream' `
            -InFile      $zipPath | Out-Null
        Write-Info "Upload complete."

        # ── g. Submit import job ──────────────────────────────────────────────
        Write-Step "Submitting import job (legal entity: $LegalEntityId)"
        $importResp = Invoke-RestMethod -Method Post `
            -Uri         "$dmfBase.ImportFromPackage" `
            -Headers     $authHeaders `
            -ContentType 'application/json' `
            -Body        (@{
                packageUrl        = $blobUrl
                definitionGroupId = $definitionGroupId
                executionId       = ''
                execute           = $true
                overwrite         = $true
                legalEntityId     = $LegalEntityId
            } | ConvertTo-Json)

        $executionId = $importResp.value
        Write-Info "Execution ID: $executionId"

        if ([string]::IsNullOrWhiteSpace($executionId)) {
            throw "ImportFromPackage did not return an execution ID. Response: $($importResp | ConvertTo-Json -Depth 5)"
        }

        # ── h. Poll for completion ────────────────────────────────────────────
        Write-Step "Waiting for completion (poll every $PollIntervalSeconds s, timeout $TimeoutMinutes min)"
        $terminalStatuses = @('Succeeded', 'PartiallySucceeded', 'Failed', 'Canceled')
        $statusBody       = @{ executionId = $executionId } | ConvertTo-Json
        $deadline         = (Get-Date).AddMinutes($TimeoutMinutes)
        $lastStatus       = ''

        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Seconds $PollIntervalSeconds
            $statusResp    = Invoke-RestMethod -Method Post `
                -Uri         "$dmfBase.GetExecutionSummaryStatus" `
                -Headers     $authHeaders `
                -ContentType 'application/json' `
                -Body        $statusBody
            $currentStatus = $statusResp.value
            if ($currentStatus -ne $lastStatus) {
                Write-Info "[$( Get-Date -Format 'HH:mm:ss')] Status: $currentStatus"
                $lastStatus = $currentStatus
            }
            if ($terminalStatuses -contains $currentStatus) { break }
        }

        # ── i. Package outcome ────────────────────────────────────────────────
        $jobHistoryPath    = "Data management > Job history > execution ID: $executionId"
        $pkgResult.ExecutionId = $executionId
        $pkgResult.Status      = if ($lastStatus) { $lastStatus } else { 'TimedOut' }

        switch ($lastStatus) {
            'Succeeded' {
                Write-OK "Import completed successfully."
            }
            'PartiallySucceeded' {
                Write-Warn "Import partially succeeded - some records may have been skipped or errored."
                Write-Info "Review: $jobHistoryPath"
            }
            'Failed' {
                Write-Fail "Import failed."
                Write-Info "Review: $jobHistoryPath"
            }
            'Canceled' {
                Write-Fail "Import was canceled."
                Write-Info "Execution ID: $executionId"
            }
            default {
                Write-Warn "Timed out after $TimeoutMinutes minutes. Last known status: '$lastStatus'"
                Write-Info "Execution ID: $executionId"
            }
        }
    }
    catch {
        $pkgResult.Status = 'Error'
        Write-Fail "Package '$pkgName' error: $_"
    }

    $results.Add($pkgResult)
}

# ─── 5. Summary ───────────────────────────────────────────────────────────────
Write-Host "`n$('─' * 72)" -ForegroundColor DarkGray
Write-Step "Summary ($($results.Count) package(s))"

foreach ($r in $results) {
    $color = switch ($r.Status) {
        'Succeeded'          { 'Green'  }
        'PartiallySucceeded' { 'Yellow' }
        default              { 'Red'    }
    }
    Write-Host ("  {0,-20}  {1,-12}  {2}" -f $r.ExecutionId, $r.Status, $r.Package) -ForegroundColor $color
}

$anyFailed = $results | Where-Object { $_.Status -in @('Failed', 'Error') }
if ($anyFailed) { exit 1 }
