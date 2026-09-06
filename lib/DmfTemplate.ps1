<#
.SYNOPSIS
    Template (Manifest.xml) reading, validation, and writing for D365 F&O DMF.

.DESCRIPTION
    A template is a folder under resources/ that contains a Manifest.xml in the
    DMF package-manifest format.  If .xlsx files are present alongside, the
    same folder is also an importable package.  Dot-source this file to import:

        Get-TemplateFolders          -- folders under a root that hold a Manifest.xml
        Get-TemplateInfo             -- metadata object for one template folder
        Read-DmfManifest             -- parse a manifest into line objects (defaults applied)
        Test-DmfManifest             -- structural warnings for a manifest file
        ConvertTo-DmfTemplateLine    -- manifest lines -> the shape DefinitionGroupTemplateLines uses
        New-DmfManifestDocument      -- build an XmlDocument from line objects
        Write-DmfManifest            -- write a manifest as UTF-16 LE with BOM (D365 requirement)
        Write-DmfPackageHeader       -- write the PackageHeader.xml boilerplate
        Read-TemplateSidecar         -- optional template.json next to the manifest
        Write-TemplateSidecar
        Get-TemplateOrigin           -- 'custom' | 'captured' | 'unknown' from template.json
        ConvertTo-DmfSafeFileName    -- strip characters D365 strips from file names

    Minimal hand-authored manifest
    ──────────────────────────────
    Only EntityName and the three ordering values are required per line;
    everything else takes a default (see Read-DmfManifest).  TargetEntity (the
    AOT entity name) is optional and is filled in by the entity resolver
    (DmfMetadata.ps1) when the OData path needs it.

    Custom templates
    ────────────────
    A template.json with "origin": "custom" marks the folder as maintained by
    hand in the repository: menus tag it [custom] and
    Export-TemplateDefinition.ps1 never overwrites it, not even with -Force.
    Folders the capture script writes carry "origin": "captured" together
    with capturedFrom / capturedAt; a folder without a sidecar (for example
    an expanded export dropped under resources/) has origin 'unknown'.

.NOTES
    Write-Warn / Write-Detail come from DmfOutput.ps1 -- dot-source it first.
    $Script:DmNs is set here if the calling script has not already set it.
#>

$Script:DmfDataManagementNs = 'http://schemas.microsoft.com/dynamics/2015/01/DataManagement'
if ($null -eq (Get-Variable -Name 'DmNs' -Scope Script -ValueOnly -ErrorAction SilentlyContinue)) {
    $Script:DmNs = $Script:DmfDataManagementNs
}


function ConvertTo-DmfSafeFileName {
    <#
    .SYNOPSIS
        Removes characters that are invalid in a file name.
    .DESCRIPTION
        Mirrors how D365 names package files: invalid characters are dropped,
        not replaced ('Country/regions' -> 'Countryregions').  Leading and
        trailing whitespace and dots are trimmed.  An empty result becomes '_'.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Name)

    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $sb = [System.Text.StringBuilder]::new()
    foreach ($ch in $Name.ToCharArray()) {
        if ($invalid -notcontains $ch) { [void]$sb.Append($ch) }
    }
    $safe = $sb.ToString().Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($safe)) { return '_' }
    return $safe
}


function Get-TemplateFolders {
    <#
    .SYNOPSIS
        Returns every sub-folder of -ResourcesPath that contains a Manifest.xml.
    .OUTPUTS
        [System.IO.DirectoryInfo] objects sorted by name.  Wrap the call in @()
        to get an array regardless of count (standard PowerShell convention;
        every array-returning function in lib/ follows it).
    #>
    param([Parameter(Mandatory)][string]$ResourcesPath)

    if (-not (Test-Path -LiteralPath $ResourcesPath -PathType Container)) { return }
    Get-ChildItem -LiteralPath $ResourcesPath -Directory |
        Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'Manifest.xml') -PathType Leaf } |
        Sort-Object -Property Name
}


function ConvertTo-DmfBoolean {
    param([AllowNull()][AllowEmptyString()]$Value, [bool]$Default)
    if ($null -eq $Value) { return $Default }
    $s = ([string]$Value).Trim().ToLowerInvariant()
    switch ($s) {
        ''      { return $Default }
        'true'  { return $true }
        'yes'   { return $true }
        '1'     { return $true }
        'false' { return $false }
        'no'    { return $false }
        '0'     { return $false }
        default { return $Default }
    }
}


function Get-DmfManifestXml {
    <#
    .SYNOPSIS  Loads a manifest file into an XmlDocument, BOM-aware (UTF-16 or UTF-8).
    #>
    param([Parameter(Mandatory)][string]$Path)

    # A StreamReader with BOM detection copes with both UTF-16 LE (written by
    # D365 and this toolset) and UTF-8 (hand-authored); LoadXml on the decoded
    # string ignores any encoding declaration that disagrees with the bytes.
    $reader = New-Object System.IO.StreamReader($Path, [System.Text.Encoding]::UTF8, $true)
    try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $false
    $doc.LoadXml($text)
    return $doc
}


function Get-DmfChildText {
    param([Parameter(Mandatory)][System.Xml.XmlNode]$Node, [Parameter(Mandatory)][string]$LocalName)
    $child = $Node.SelectSingleNode("*[local-name()='$LocalName']")
    if ($null -eq $child) { return $null }
    return $child.InnerText
}


function Read-DmfManifest {
    <#
    .SYNOPSIS
        Parses a Manifest.xml into a template object with one line per entity.

    .DESCRIPTION
        Applies defaults so callers can rely on every property being present:
        ordering values 1, FailLevelOnError / FailExecutionUnitOnError /
        Disable false, RunBusinessLogic / RunBusinessValidation true,
        SourceFormat EXCEL, InputFilePath '<EntityName>.xlsx' and
        ExcelSheetName '<Entity_Name>$' (invalid file-name characters removed,
        spaces to underscores).  The raw XML element of
        each line is kept (RawNode) so New-DmfManifestDocument can re-emit a
        D365-generated line verbatim (field maps, query data) while still
        overriding the ordering values.

    .OUTPUTS
        [pscustomobject]  Path, DefinitionGroupName, Description, Lines[], Document
        Each line: EntityName, TargetEntity, ExecutionUnit, LevelInExecutionUnit,
        SequenceInLevel, FailLevelOnError, FailExecutionUnitOnError,
        RunBusinessLogic, RunBusinessValidation, Disable, SourceFormat,
        InputFilePath, ExcelSheetName, RawNode
    #>
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Manifest not found: '$Path'" }
    $doc  = Get-DmfManifestXml -Path $Path
    $root = $doc.DocumentElement
    if ($null -eq $root -or $root.LocalName -ne 'DataManagementPackageManifest') {
        throw "'$Path' is not a DMF manifest (root element is '$($root.LocalName)', expected DataManagementPackageManifest)."
    }

    $lines = [System.Collections.Generic.List[pscustomobject]]::new()
    foreach ($node in $root.SelectNodes("*[local-name()='PackageEntityList']/*[local-name()='DataManagementPackageEntityData']")) {
        $entityName = Get-DmfChildText -Node $node -LocalName 'EntityName'
        if ([string]::IsNullOrWhiteSpace($entityName)) { continue }
        $safeName   = ConvertTo-DmfSafeFileName -Name $entityName

        $toInt = {
            param($raw, $default)
            $v = 0
            if ($null -ne $raw -and [int]::TryParse(([string]$raw).Trim(), [ref]$v)) { return $v }
            return $default
        }

        $inputFile = Get-DmfChildText -Node $node -LocalName 'InputFilePath'
        if ([string]::IsNullOrWhiteSpace($inputFile)) { $inputFile = "$safeName.xlsx" }
        $sheet = Get-DmfChildText -Node $node -LocalName 'ExcelSheetName'
        if ([string]::IsNullOrWhiteSpace($sheet)) { $sheet = ($safeName -replace '\s+', '_') + '$' }
        $sourceFormat = Get-DmfChildText -Node $node -LocalName 'SourceFormat'
        if ([string]::IsNullOrWhiteSpace($sourceFormat)) { $sourceFormat = 'EXCEL' }
        $target = Get-DmfChildText -Node $node -LocalName 'TargetEntity'
        if ([string]::IsNullOrWhiteSpace($target)) { $target = $null } else { $target = $target.Trim() }

        $lines.Add([pscustomobject]@{
            EntityName               = $entityName.Trim()
            TargetEntity             = $target
            ExecutionUnit            = & $toInt (Get-DmfChildText -Node $node -LocalName 'ExecutionUnit') 1
            LevelInExecutionUnit     = & $toInt (Get-DmfChildText -Node $node -LocalName 'LevelInExecutionUnit') 1
            SequenceInLevel          = & $toInt (Get-DmfChildText -Node $node -LocalName 'SequenceInLevel') 1
            FailLevelOnError         = ConvertTo-DmfBoolean (Get-DmfChildText -Node $node -LocalName 'FailLevelOnError') $false
            FailExecutionUnitOnError = ConvertTo-DmfBoolean (Get-DmfChildText -Node $node -LocalName 'FailExecutionUnitOnError') $false
            RunBusinessLogic         = ConvertTo-DmfBoolean (Get-DmfChildText -Node $node -LocalName 'RunBusinessLogic') $true
            RunBusinessValidation    = ConvertTo-DmfBoolean (Get-DmfChildText -Node $node -LocalName 'RunBusinessValidation') $true
            Disable                  = ConvertTo-DmfBoolean (Get-DmfChildText -Node $node -LocalName 'Disable') $false
            SourceFormat             = $sourceFormat
            InputFilePath            = $inputFile
            ExcelSheetName           = $sheet
            RawNode                  = $node
        })
    }

    $dgn  = Get-DmfChildText -Node $root -LocalName 'DefinitionGroupName'
    $desc = Get-DmfChildText -Node $root -LocalName 'Description'

    return [pscustomobject]@{
        Path                = $Path
        DefinitionGroupName = if ($dgn)  { $dgn }  else { '' }
        Description         = if ($desc) { $desc } else { '' }
        Lines               = $lines.ToArray()
        Document            = $doc
    }
}


function Test-DmfManifest {
    <#
    .SYNOPSIS
        Returns structural warnings for a manifest file (nothing = clean).
    .OUTPUTS
        [string] warnings; wrap the call in @() to count them.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $warnings = [System.Collections.Generic.List[string]]::new()
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $warnings.Add("Manifest.xml not found at '$Path'")
        return $warnings.ToArray()
    }

    $doc = $null
    try { $doc = Get-DmfManifestXml -Path $Path }
    catch {
        $warnings.Add("Manifest.xml could not be parsed: $($_.Exception.Message)")
        return $warnings.ToArray()
    }

    $root = $doc.DocumentElement
    if ($null -eq $root -or $root.LocalName -ne 'DataManagementPackageManifest') {
        $warnings.Add("Root element is '$($root.LocalName)'; expected DataManagementPackageManifest")
        return $warnings.ToArray()
    }
    if ($root.NamespaceURI -ne $Script:DmfDataManagementNs) {
        $warnings.Add("Root namespace is '$($root.NamespaceURI)'; expected $Script:DmfDataManagementNs")
    }

    $nodes = $root.SelectNodes("*[local-name()='PackageEntityList']/*[local-name()='DataManagementPackageEntityData']")
    if ($nodes.Count -eq 0) {
        $warnings.Add('No entity lines (PackageEntityList/DataManagementPackageEntityData) found')
        return $warnings.ToArray()
    }

    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $i = 0
    foreach ($node in $nodes) {
        $i++
        $name = Get-DmfChildText -Node $node -LocalName 'EntityName'
        if ([string]::IsNullOrWhiteSpace($name)) {
            $warnings.Add("Line $i has no EntityName")
            continue
        }
        if (-not $seen.Add($name.Trim())) { $warnings.Add("Duplicate EntityName '$($name.Trim())'") }

        foreach ($ordName in 'ExecutionUnit', 'LevelInExecutionUnit', 'SequenceInLevel') {
            $raw = Get-DmfChildText -Node $node -LocalName $ordName
            if ($null -eq $raw) { $warnings.Add("'$name': $ordName is missing"); continue }
            $tmp = 0
            if (-not [int]::TryParse($raw.Trim(), [ref]$tmp)) { $warnings.Add("'$name': $ordName '$raw' is not an integer") }
        }
        foreach ($boolName in 'Disable', 'FailLevelOnError', 'FailExecutionUnitOnError', 'RunBusinessLogic', 'RunBusinessValidation') {
            $raw = Get-DmfChildText -Node $node -LocalName $boolName
            if ($null -ne $raw -and $raw.Trim() -ne '' -and $raw.Trim().ToLowerInvariant() -notin 'true', 'false', 'yes', 'no', '0', '1') {
                $warnings.Add("'$name': $boolName '$raw' is not a boolean")
            }
        }
    }
    return $warnings.ToArray()
}


function ConvertTo-DmfTemplateLine {
    <#
    .SYNOPSIS
        Converts manifest lines to the property shape of DefinitionGroupTemplateLines.
    .DESCRIPTION
        Invoke-ProjectExport.ps1 posts project lines from either source; this
        lets a local manifest feed the same code path as an environment
        template.  Output properties: Entity, ExecutionUnit,
        LevelInExecutionUnit, Sequence, FailLevelOnError ('Yes'/'No'),
        FailExecutionUnitOnError ('Yes'/'No'), TargetEntity, Disable.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Lines,
        [switch]$IncludeDisabled
    )
    $out = [System.Collections.Generic.List[pscustomobject]]::new()
    foreach ($l in $Lines) {
        if ($l.Disable -and -not $IncludeDisabled) { continue }
        $out.Add([pscustomobject]@{
            Entity                   = $l.EntityName
            TargetEntity             = $l.TargetEntity
            ExecutionUnit            = [int]$l.ExecutionUnit
            LevelInExecutionUnit     = [int]$l.LevelInExecutionUnit
            Sequence                 = [int]$l.SequenceInLevel
            FailLevelOnError         = $(if ($l.FailLevelOnError) { 'Yes' } else { 'No' })
            FailExecutionUnitOnError = $(if ($l.FailExecutionUnitOnError) { 'Yes' } else { 'No' })
            Disable                  = [bool]$l.Disable
        })
    }
    return $out.ToArray()
}


function New-DmfManifestDocument {
    <#
    .SYNOPSIS
        Builds a manifest XmlDocument from template line objects.

    .DESCRIPTION
        For a line that carries a RawNode (read from an existing manifest) the
        element is imported verbatim -- preserving EntityMapList, QueryData and
        the rest -- and only ExecutionUnit / LevelInExecutionUnit /
        SequenceInLevel (and TargetEntity when the node lacks one) are set
        from the line.  Lines without a RawNode are written from their
        properties in the element order D365 uses.

    .PARAMETER Lines
        Objects as returned by Read-DmfManifest (RawNode optional).  Objects
        with only EntityName / ordering properties are accepted; other
        properties default as in Read-DmfManifest.

    .OUTPUTS
        [System.Xml.XmlDocument]
    #>
    param(
        [Parameter(Mandatory)][string]$DefinitionGroupName,
        [AllowEmptyString()][string]$Description = '',
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Lines
    )

    $ns  = $Script:DmfDataManagementNs
    $doc = New-Object System.Xml.XmlDocument
    [void]$doc.AppendChild($doc.CreateXmlDeclaration('1.0', 'utf-16', $null))
    $root = $doc.CreateElement('DataManagementPackageManifest', $ns)
    [void]$root.SetAttribute('xmlns:i', 'http://www.w3.org/2001/XMLSchema-instance')
    [void]$doc.AppendChild($root)

    $addText = {
        param($parent, $name, $value)
        $el = $doc.CreateElement($name, $ns)
        $el.InnerText = [string]$value
        [void]$parent.AppendChild($el)
        return $el
    }

    [void](& $addText $root 'DefinitionGroupName' $DefinitionGroupName)
    [void](& $addText $root 'Description' $Description)
    $listEl = $doc.CreateElement('PackageEntityList', $ns)
    [void]$root.AppendChild($listEl)

    $prop = {
        param($obj, $name, $default)
        $p = $obj.PSObject.Properties[$name]
        if ($null -ne $p -and $null -ne $p.Value) { return $p.Value }
        return $default
    }
    $boolText = { param($b) if ($b) { 'true' } else { 'false' } }

    foreach ($line in $Lines) {
        $entityName = [string](& $prop $line 'EntityName' '')
        if ([string]::IsNullOrWhiteSpace($entityName)) { continue }
        $safeName = ConvertTo-DmfSafeFileName -Name $entityName

        $eu  = [int](& $prop $line 'ExecutionUnit' 1)
        $lv  = [int](& $prop $line 'LevelInExecutionUnit' 1)
        $seq = [int](& $prop $line 'SequenceInLevel' 1)
        $target = & $prop $line 'TargetEntity' $null

        $raw = & $prop $line 'RawNode' $null
        if ($null -ne $raw -and $raw -is [System.Xml.XmlNode]) {
            $imported = $doc.ImportNode($raw, $true)
            foreach ($pair in @(@('ExecutionUnit', $eu), @('LevelInExecutionUnit', $lv), @('SequenceInLevel', $seq))) {
                $el = $imported.SelectSingleNode("*[local-name()='$($pair[0])']")
                if ($null -eq $el) { $el = $doc.CreateElement($pair[0], $ns); [void]$imported.AppendChild($el) }
                $el.InnerText = [string]$pair[1]
            }
            if ($target) {
                $tEl = $imported.SelectSingleNode("*[local-name()='TargetEntity']")
                if ($null -eq $tEl) { $tEl = $doc.CreateElement('TargetEntity', $ns); [void]$imported.AppendChild($tEl) }
                if ([string]::IsNullOrWhiteSpace($tEl.InnerText)) { $tEl.InnerText = [string]$target }
            }
            [void]$listEl.AppendChild($imported)
            continue
        }

        # -- Build from properties, in D365's element order -------------------
        $entEl = $doc.CreateElement('DataManagementPackageEntityData', $ns)
        [void](& $addText $entEl 'Disable' (& $boolText (& $prop $line 'Disable' $false)))
        # The three-argument SetAttribute overload returns the value; void it
        # so it does not leak into the function's output stream.
        $mapList = $doc.CreateElement('EntityMapList', $ns)
        [void]$mapList.SetAttribute('nil', 'http://www.w3.org/2001/XMLSchema-instance', 'true')
        [void]$entEl.AppendChild($mapList)
        [void](& $addText $entEl 'EntityName' $entityName)
        $xfList = $doc.CreateElement('EntityTransformList', $ns)
        [void]$xfList.SetAttribute('nil', 'http://www.w3.org/2001/XMLSchema-instance', 'true')
        [void]$entEl.AppendChild($xfList)
        $sheet = & $prop $line 'ExcelSheetName' (($safeName -replace '\s+', '_') + '$')
        [void](& $addText $entEl 'ExcelSheetName' $sheet)
        [void](& $addText $entEl 'ExecutionUnit' $eu)
        [void](& $addText $entEl 'FailExecutionUnitOnError' (& $boolText (& $prop $line 'FailExecutionUnitOnError' $false)))
        [void](& $addText $entEl 'FailLevelOnError' (& $boolText (& $prop $line 'FailLevelOnError' $false)))
        [void](& $addText $entEl 'InputFilePath' (& $prop $line 'InputFilePath' "$safeName.xlsx"))
        [void](& $addText $entEl 'LevelInExecutionUnit' $lv)
        [void](& $addText $entEl 'QueryFilter' '')
        [void](& $addText $entEl 'RunBusinessLogic' (& $boolText (& $prop $line 'RunBusinessLogic' $true)))
        [void](& $addText $entEl 'RunBusinessValidation' (& $boolText (& $prop $line 'RunBusinessValidation' $true)))
        [void](& $addText $entEl 'SequenceInLevel' $seq)
        [void](& $addText $entEl 'SourceFormat' (& $prop $line 'SourceFormat' 'EXCEL'))
        if ($target) { [void](& $addText $entEl 'TargetEntity' $target) }
        [void]$listEl.AppendChild($entEl)
    }

    [void](& $addText $root 'ProjectCategory' '1')
    [void](& $addText $root 'RulesData' '')
    return $doc
}


function Write-DmfManifest {
    <#
    .SYNOPSIS
        Saves a manifest XmlDocument as UTF-16 LE with BOM (required by D365 DMF).
    #>
    param(
        [Parameter(Mandatory)][System.Xml.XmlDocument]$Document,
        [Parameter(Mandatory)][string]$Path
    )
    $settings          = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = [System.Text.Encoding]::Unicode
    $settings.Indent   = $true
    $writer = [System.Xml.XmlWriter]::Create($Path, $settings)
    try     { $Document.Save($writer) }
    finally { $writer.Dispose() }
}


function Write-DmfPackageHeader {
    <#
    .SYNOPSIS
        Writes PackageHeader.xml (UTF-16 LE with BOM) with the DMF boilerplate.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [AllowEmptyString()][string]$Description = ''
    )
    $descXml   = [System.Security.SecurityElement]::Escape($Description)
    $headerXml = (@(
        '<?xml version="1.0" encoding="utf-16"?>',
        '<DataManagementPackageHeader xmlns:i="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/dynamics/2015/01/DataManagement">',
        "  <Description>$descXml</Description>",
        '  <ManifestType>Microsoft.Dynamics.AX.Framework.Tools.DataManagement.Serialization.DataManagementPackageManifest</ManifestType>',
        '  <PackageType>DefinitionGroup</PackageType>',
        '  <PackageVersion>2</PackageVersion>',
        '</DataManagementPackageHeader>'
    ) -join [System.Environment]::NewLine)
    [System.IO.File]::WriteAllText($Path, $headerXml, [System.Text.Encoding]::Unicode)
}


function Read-TemplateSidecar {
    <#
    .SYNOPSIS  Reads template.json from a template folder; $null when absent or unreadable.
    #>
    param([Parameter(Mandatory)][string]$Folder)
    $path = Join-Path $Folder 'template.json'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    try { return (Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json) }
    catch {
        Write-Warn "template.json in '$Folder' could not be parsed: $($_.Exception.Message)"
        return $null
    }
}


function Write-TemplateSidecar {
    <#
    .SYNOPSIS  Writes template.json (UTF-8, pretty-printed) into a template folder.
    #>
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)]$Sidecar
    )
    $path = Join-Path $Folder 'template.json'
    $json = $Sidecar | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
}


function Get-TemplateOrigin {
    <#
    .SYNOPSIS
        Classifies a template folder from its template.json.

    .DESCRIPTION
        Returns 'custom' when the sidecar has "origin": "custom" (hand-authored
        and maintained in the repository; capture never overwrites it),
        'captured' when it has "origin": "captured" or a capturedFrom value
        (a mirror of an environment template that re-capture may refresh), and
        'unknown' when there is no readable sidecar or it says neither.
        The origin value is matched case-insensitively.

    .PARAMETER Folder
        Template folder path.

    .PARAMETER Sidecar
        An already-read sidecar (object or hashtable); skips the file read.
    #>
    param(
        [Parameter(Mandatory)][string]$Folder,
        $Sidecar
    )
    if ($null -eq $Sidecar) { $Sidecar = Read-TemplateSidecar -Folder $Folder }
    if ($null -eq $Sidecar) { return 'unknown' }

    $get = {
        param($obj, $name)
        if ($obj -is [System.Collections.IDictionary]) {
            foreach ($k in $obj.Keys) { if ([string]$k -eq $name) { return $obj[$k] } }
            return $null
        }
        $p = $obj.PSObject.Properties[$name]
        if ($null -ne $p) { return $p.Value } else { return $null }
    }

    $origin = [string](& $get $Sidecar 'origin')
    if ($origin) { $origin = $origin.Trim().ToLowerInvariant() }
    if ($origin -eq 'custom')   { return 'custom' }
    if ($origin -eq 'captured') { return 'captured' }
    if ([string](& $get $Sidecar 'capturedFrom')) { return 'captured' }
    return 'unknown'
}


function Get-TemplateInfo {
    <#
    .SYNOPSIS
        Inspects a template folder and returns a metadata object for menus.

    .PARAMETER EntityMap
        Optional entity map (DmfMetadata.ps1) used to count lines that already
        resolve to an OData collection without any network access.

    .OUTPUTS
        [pscustomobject]  Index, Folder, Name, Description, DefinitionGroupName,
        EntityCount, DisabledCount, HasData, XlsxCount, HasSidecar, Origin,
        IsCustom, HasOrdering, ResolvedCount, IsValid, Warnings, Lines
    #>
    param(
        [Parameter(Mandatory)][System.IO.DirectoryInfo]$Folder,
        [Parameter(Mandatory)][int]$Index,
        $EntityMap
    )

    $manifestPath = Join-Path $Folder.FullName 'Manifest.xml'
    $warnings     = [System.Collections.Generic.List[string]]::new()
    foreach ($w in (Test-DmfManifest -Path $manifestPath)) { $warnings.Add($w) }

    $lines = @(); $desc = ''; $dgn = ''
    try {
        $m     = Read-DmfManifest -Path $manifestPath
        $lines = $m.Lines
        $desc  = $m.Description
        $dgn   = $m.DefinitionGroupName
    } catch {
        if ($warnings.Count -eq 0) { $warnings.Add("Manifest.xml could not be read: $($_.Exception.Message)") }
    }

    $xlsx = @(Get-ChildItem -LiteralPath $Folder.FullName -Filter '*.xlsx' -File -ErrorAction SilentlyContinue)

    $sidecar = Read-TemplateSidecar -Folder $Folder.FullName
    $origin  = Get-TemplateOrigin -Folder $Folder.FullName -Sidecar $sidecar

    $resolved = $null
    if ($null -ne $EntityMap -and $lines.Count -gt 0 -and (Get-Command -Name 'Test-DmfEntityMapResolved' -ErrorAction SilentlyContinue)) {
        $resolved = @($lines | Where-Object { Test-DmfEntityMapResolved -Map $EntityMap -EntityName $_.EntityName }).Count
    }

    return [pscustomobject]@{
        Index               = $Index
        Folder              = $Folder
        Name                = $Folder.Name
        Description         = $desc
        DefinitionGroupName = $dgn
        EntityCount         = $lines.Count
        DisabledCount       = @($lines | Where-Object { $_.Disable }).Count
        HasData             = ($xlsx.Count -gt 0)
        XlsxCount           = $xlsx.Count
        HasSidecar          = (Test-Path -LiteralPath (Join-Path $Folder.FullName 'template.json') -PathType Leaf)
        Origin              = $origin
        IsCustom            = ($origin -eq 'custom')
        HasOrdering         = (Test-Path -LiteralPath (Join-Path $Folder.FullName 'ordering.json') -PathType Leaf)
        ResolvedCount       = $resolved
        IsValid             = ($lines.Count -gt 0 -and @($warnings | Where-Object { $_ -notlike 'Root namespace*' }).Count -eq 0)
        Warnings            = $warnings.ToArray()
        Lines               = $lines
    }
}
