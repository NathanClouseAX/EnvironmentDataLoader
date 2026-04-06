<#
.SYNOPSIS
    Package discovery and entity execution ordering for D365 F&O DMF imports.

.DESCRIPTION
    Dot-source this file to import the following into your script:
        $entityOrdering         -- ordered hashtable of built-in EU/LV/SEQ defaults
        Get-PackageInfo         -- inspects a package folder and returns metadata
        Resolve-EntityOrdering  -- merges built-in defaults with per-package overrides

    Per-package overrides
    ─────────────────────
    Place an ordering.json file in any package folder to customise entity
    execution order for that package.  Entries are merged with the built-in
    defaults; package-level values take precedence for any named entity.

    ordering.json format:
    {
      "Entity Name": { "EU": 1, "LV": 10, "SEQ": 20 },
      ...
    }

.NOTES
    Write-Warn and Write-Detail are called by Resolve-EntityOrdering -- ensure
    DmfOutput.ps1 is dot-sourced before calling these functions.

    The $Script:DmNs variable (D365 XML namespace URI) must be set by the caller
    before using Get-PackageInfo.
#>

# =============================================================================
#  Entity execution ordering  (built-in defaults)
#
#  Entities sharing the same EU + LV + SEQ are processed in parallel.
#  Within an EU, levels run sequentially (ascending).
#  Within a level, sequences run sequentially (ascending).
#
#  EU=1  LV=10  Foundation + geographic + org data (prerequisite for all below)
#  EU=1  LV=20  Finance (GL) and Workflow merged into one level so their
#               independent chains advance in parallel at each matching SEQ step.
#
#  Entities absent from this table fall back to their source Manifest.xml values.
#  Use a per-package ordering.json to extend or override for new datasets.
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
#  Functions
# =============================================================================

function Get-PackageInfo {
    <#
    .SYNOPSIS
        Inspects a package folder and returns a metadata object.

    .PARAMETER Folder
        The DirectoryInfo object representing the package folder.

    .PARAMETER Index
        One-based display index for the package in selection menus.

    .OUTPUTS
        [pscustomobject] with properties:
            Index, Folder, Name, XlsxCount, XlsxSizeMB, EntityCount,
            HasManifest, HasOrdering, IsValid, Warnings
    #>
    param(
        [Parameter(Mandatory)] [System.IO.DirectoryInfo]$Folder,
        [Parameter(Mandatory)] [int]$Index
    )

    $xlsxCount   = 0
    $xlsxSizeMB  = 0
    $entityCount = 0
    $hasManifest = $false
    $hasOrdering = $false
    $warnings    = [System.Collections.Generic.List[string]]::new()

    try {
        $xlsxFiles = @(Get-ChildItem -Path $Folder.FullName -Filter '*.xlsx' -File -ErrorAction Stop)
        $xlsxCount = $xlsxFiles.Count
        if ($xlsxCount -gt 0) {
            $xlsxSizeMB = [Math]::Round(
                ($xlsxFiles | Measure-Object -Property Length -Sum).Sum / 1MB, 2)
        }
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

    $hasOrdering = Test-Path (Join-Path $Folder.FullName 'ordering.json') -PathType Leaf

    return [pscustomobject]@{
        Index       = $Index
        Folder      = $Folder
        Name        = $Folder.Name
        XlsxCount   = $xlsxCount
        XlsxSizeMB  = $xlsxSizeMB
        EntityCount = $entityCount
        HasManifest = $hasManifest
        HasOrdering = $hasOrdering
        IsValid     = ($hasManifest -and $xlsxCount -gt 0)
        Warnings    = $warnings
    }
}

function Resolve-EntityOrdering {
    <#
    .SYNOPSIS
        Returns the effective entity ordering for a package.

    .DESCRIPTION
        Clones the built-in $entityOrdering defaults then applies any overrides
        found in the package folder's ordering.json.  Package-level entries
        take precedence for any entity name present in both sources.

    .PARAMETER PackageFolderPath
        Full path to the package folder (may contain ordering.json).

    .OUTPUTS
        [ordered hashtable]  keys = entity name, values = @{ EU; LV; SEQ }
    #>
    param([Parameter(Mandatory)][string]$PackageFolderPath)

    # Shallow-clone the built-in defaults
    $merged = [ordered]@{}
    foreach ($key in $entityOrdering.Keys) { $merged[$key] = $entityOrdering[$key] }

    # Apply per-package overrides when ordering.json is present
    $localPath = Join-Path $PackageFolderPath 'ordering.json'
    if (Test-Path $localPath -PathType Leaf) {
        try {
            $localOrdering = Get-Content -Path $localPath -Raw -Encoding UTF8 | ConvertFrom-Json
            $overrideCount = 0
            foreach ($prop in $localOrdering.PSObject.Properties) {
                $merged[$prop.Name] = @{
                    EU  = [int]$prop.Value.EU
                    LV  = [int]$prop.Value.LV
                    SEQ = [int]$prop.Value.SEQ
                }
                $overrideCount++
            }
            Write-Detail "ordering.json : $overrideCount override(s) merged."
        } catch {
            Write-Warn "Could not parse ordering.json: $_ -- using built-in defaults."
        }
    }

    return $merged
}
