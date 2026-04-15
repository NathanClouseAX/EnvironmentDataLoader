<#
.SYNOPSIS
    Zip inspection helpers for D365 F&O DMF package tooling.

.DESCRIPTION
    Dot-source this file to import Get-ZipPackageInfo into your script.

    The following $Script: variable must be set by the caller:
        $Script:DmNs  -- D365 XML namespace URI

    Requires System.IO.Compression.FileSystem to be loaded by the caller:
        Add-Type -AssemblyName System.IO.Compression.FileSystem
#>

function Get-ZipPackageInfo {
    <#
    .SYNOPSIS
        Inspects a DMF package .zip and returns metadata from its Manifest.xml.

    .PARAMETER File
        The FileInfo object for the .zip file.

    .PARAMETER Index
        One-based display index for selection menus.

    .OUTPUTS
        [pscustomobject] with properties:
            Index, File, Name, SizeMB, DefinitionGroupId, EntityCount,
            HasManifest, IsValid, Warnings
    #>
    param(
        [Parameter(Mandatory)] [System.IO.FileInfo]$File,
        [Parameter(Mandatory)] [int]$Index
    )

    $sizeMB          = [Math]::Round($File.Length / 1MB, 2)
    $definitionGroup = ''
    $entityCount     = 0
    $hasManifest     = $false
    $warnings        = [System.Collections.Generic.List[string]]::new()

    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($File.FullName)
        try {
            $manifestEntry = $zip.Entries |
                Where-Object { $_.Name -eq 'Manifest.xml' } |
                Select-Object -First 1

            if ($manifestEntry) {
                $hasManifest = $true
                try {
                    $stream  = $manifestEntry.Open()
                    # BOM-aware reader: handles both UTF-16 LE (written by this toolset)
                    # and UTF-8 (written by D365 exports).
                    $reader  = New-Object System.IO.StreamReader($stream, $true)
                    $xmlText = $reader.ReadToEnd()
                    $reader.Dispose()
                    $stream.Dispose()

                    $doc   = New-Object System.Xml.XmlDocument
                    $doc.LoadXml($xmlText)
                    $nsMgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
                    $nsMgr.AddNamespace('dm', $Script:DmNs)

                    $dgNode = $doc.SelectSingleNode('//dm:DefinitionGroupName', $nsMgr)
                    if ($dgNode) { $definitionGroup = $dgNode.InnerText }
                    else         { $warnings.Add('DefinitionGroupName not found in Manifest.xml') }

                    $entityCount = $doc.SelectNodes(
                        '//*[local-name()="DataManagementPackageEntityData"]').Count
                } catch {
                    $warnings.Add("Manifest.xml could not be parsed: $_")
                }
            } else {
                $warnings.Add('No Manifest.xml found inside zip')
            }
        } finally {
            $zip.Dispose()
        }
    } catch {
        $warnings.Add("Could not open zip: $_")
    }

    return [pscustomobject]@{
        Index             = $Index
        File              = $File
        Name              = $File.Name
        SizeMB            = $sizeMB
        DefinitionGroupId = $definitionGroup
        EntityCount       = $entityCount
        HasManifest       = $hasManifest
        IsValid           = ($hasManifest -and $definitionGroup -ne '')
        Warnings          = $warnings
    }
}
