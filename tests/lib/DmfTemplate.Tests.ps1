#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
    Unit tests for lib/DmfTemplate.ps1.  Uses the real resources/010 - System Setup
    manifest (92 lines, UTF-16) and small fixtures under tests/fixtures/manifests.
#>
Set-StrictMode -Version Latest

BeforeAll {
    $repoRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $libPath  = Join-Path $repoRoot 'lib'
    . (Join-Path $libPath 'DmfOutput.ps1')
    . (Join-Path $libPath 'DmfTemplate.ps1')
    $Script:LineWidth = 80
    $Script:Version   = 'test'

    $fixtures  = Join-Path $repoRoot 'tests\fixtures\manifests'
    $realPath  = Join-Path $repoRoot 'resources\010 - System Setup\Manifest.xml'
    $tempRoot  = Join-Path ([System.IO.Path]::GetTempPath()) ("DmfTemplateTests_" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
}

AfterAll {
    if (Test-Path $tempRoot) { Remove-Item $tempRoot -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'ConvertTo-DmfSafeFileName' {
    It 'removes the slash the way D365 does' {
        ConvertTo-DmfSafeFileName -Name 'Country/regions' | Should -Be 'Countryregions'
    }
    It 'keeps spaces and trims' {
        ConvertTo-DmfSafeFileName -Name '  Exchange rates. ' | Should -Be 'Exchange rates'
    }
    It 'returns _ for nothing usable' {
        ConvertTo-DmfSafeFileName -Name '???' | Should -Be '_'
    }
}

Describe 'Read-DmfManifest' {
    It 'reads the real UTF-16 D365 manifest with TargetEntity on every line' {
        $m = Read-DmfManifest -Path $realPath
        $m.Lines.Count          | Should -Be 92
        $m.DefinitionGroupName  | Should -Be 'System and Shared'
        @($m.Lines | Where-Object { -not $_.TargetEntity }).Count | Should -Be 0
        $first = $m.Lines | Where-Object EntityName -eq 'Account structure activation'
        $first.TargetEntity          | Should -Be 'LedgerAccountStructureActivationEntity'
        $first.ExecutionUnit         | Should -Be 1
        $first.LevelInExecutionUnit  | Should -Be 20
        $first.SequenceInLevel       | Should -Be 44
        $first.RawNode               | Should -Not -BeNullOrEmpty
    }

    It 'reads a minimal UTF-8 manifest and applies the defaults' {
        $m = Read-DmfManifest -Path (Join-Path $fixtures 'minimal-utf8\Manifest.xml')
        $m.Lines.Count | Should -Be 3
        $c = $m.Lines[0]
        $c.EntityName               | Should -Be 'Currencies'
        $c.TargetEntity             | Should -BeNullOrEmpty
        $c.Disable                  | Should -BeFalse
        $c.FailLevelOnError         | Should -BeFalse
        $c.RunBusinessLogic         | Should -BeTrue
        $c.SourceFormat             | Should -Be 'EXCEL'
        $c.InputFilePath            | Should -Be 'Currencies.xlsx'
        $c.ExcelSheetName           | Should -Be 'Currencies$'
        $m.Lines[1].Disable         | Should -BeTrue
        $m.Lines[1].TargetEntity    | Should -Be 'ExchangeRateEntity'
        $m.Lines[2].InputFilePath   | Should -Be 'Countryregions.xlsx'
        $m.Lines[2].ExcelSheetName  | Should -Be 'Countryregions$'
    }

    It 'rejects a non-manifest XML file' {
        $p = Join-Path $tempRoot 'notamanifest.xml'
        Set-Content -Path $p -Value '<Other/>' -Encoding UTF8
        { Read-DmfManifest -Path $p } | Should -Throw '*not a DMF manifest*'
    }
}

Describe 'Test-DmfManifest' {
    It 'is clean for the real manifest and the minimal fixture' {
        @(Test-DmfManifest -Path $realPath).Count | Should -Be 0
        @(Test-DmfManifest -Path (Join-Path $fixtures 'minimal-utf8\Manifest.xml')).Count | Should -Be 0
    }
    It 'reports each structural problem in the broken fixture' {
        $w = @(Test-DmfManifest -Path (Join-Path $fixtures 'broken\Manifest.xml'))
        $w | Should -Contain "'Currencies': ExecutionUnit 'abc' is not an integer"
        $w | Should -Contain "Duplicate EntityName 'Currencies'"
        $w | Should -Contain "'Currencies': SequenceInLevel is missing"
        $w | Should -Contain "'Currencies': Disable 'maybe' is not a boolean"
        $w | Should -Contain 'Line 3 has no EntityName'
    }
    It 'reports a missing file' {
        @(Test-DmfManifest -Path (Join-Path $tempRoot 'nope\Manifest.xml'))[0] | Should -BeLike 'Manifest.xml not found*'
    }
}

Describe 'ConvertTo-DmfTemplateLine' {
    It 'maps to the DefinitionGroupTemplateLines shape and drops disabled lines' {
        $m = Read-DmfManifest -Path (Join-Path $fixtures 'minimal-utf8\Manifest.xml')
        $t = @(ConvertTo-DmfTemplateLine -Lines $m.Lines)
        $t.Count                       | Should -Be 2
        $t[0].Entity                   | Should -Be 'Currencies'
        $t[0].Sequence                 | Should -Be 10
        $t[0].FailLevelOnError         | Should -Be 'No'
        $t[0].FailExecutionUnitOnError | Should -Be 'No'
        @(ConvertTo-DmfTemplateLine -Lines $m.Lines -IncludeDisabled).Count | Should -Be 3
    }
}

Describe 'New-DmfManifestDocument / Write-DmfManifest round trip' {
    It 'round-trips the real manifest line for line with a UTF-16 BOM' {
        $m   = Read-DmfManifest -Path $realPath
        $doc = New-DmfManifestDocument -DefinitionGroupName 'RT' -Description 'round trip' -Lines $m.Lines
        $out = Join-Path $tempRoot 'roundtrip\Manifest.xml'
        New-Item -ItemType Directory -Path (Split-Path $out) -Force | Out-Null
        Write-DmfManifest -Document $doc -Path $out

        $bytes = [System.IO.File]::ReadAllBytes($out)
        $bytes[0] | Should -Be 0xFF
        $bytes[1] | Should -Be 0xFE

        $back = Read-DmfManifest -Path $out
        $back.DefinitionGroupName | Should -Be 'RT'
        $back.Lines.Count         | Should -Be 92
        for ($i = 0; $i -lt 92; $i++) {
            $back.Lines[$i].EntityName      | Should -Be $m.Lines[$i].EntityName
            $back.Lines[$i].TargetEntity    | Should -Be $m.Lines[$i].TargetEntity
            $back.Lines[$i].SequenceInLevel | Should -Be $m.Lines[$i].SequenceInLevel
        }
        # Verbatim import keeps the field map of a D365-generated line
        $doc.SelectNodes("//*[local-name()='EntityMap']").Count | Should -BeGreaterThan 0
    }

    It 'overrides ordering on an imported node and fills a missing TargetEntity' {
        $m = Read-DmfManifest -Path (Join-Path $fixtures 'minimal-utf8\Manifest.xml')
        $m.Lines[0].ExecutionUnit = 7
        $m.Lines[0].TargetEntity  = 'CurrencyEntity'
        $doc  = New-DmfManifestDocument -DefinitionGroupName 'x' -Lines $m.Lines
        $node = $doc.SelectSingleNode("//*[local-name()='DataManagementPackageEntityData'][*[local-name()='EntityName']='Currencies']")
        $node.SelectSingleNode("*[local-name()='ExecutionUnit']").InnerText | Should -Be '7'
        $node.SelectSingleNode("*[local-name()='TargetEntity']").InnerText  | Should -Be 'CurrencyEntity'
    }

    It 'builds a line from bare properties in D365 element order' {
        $lines = @([pscustomobject]@{ EntityName = 'Units'; ExecutionUnit = 1; LevelInExecutionUnit = 10; SequenceInLevel = 20; TargetEntity = 'UnitOfMeasureEntity' })
        $doc  = New-DmfManifestDocument -DefinitionGroupName 'g' -Description 'd' -Lines $lines
        $node = $doc.SelectSingleNode("//*[local-name()='DataManagementPackageEntityData']")
        @($node.ChildNodes | ForEach-Object { $_.LocalName }) | Should -Be @(
            'Disable', 'EntityMapList', 'EntityName', 'EntityTransformList', 'ExcelSheetName', 'ExecutionUnit',
            'FailExecutionUnitOnError', 'FailLevelOnError', 'InputFilePath', 'LevelInExecutionUnit', 'QueryFilter',
            'RunBusinessLogic', 'RunBusinessValidation', 'SequenceInLevel', 'SourceFormat', 'TargetEntity')
        $node.SelectSingleNode("*[local-name()='InputFilePath']").InnerText | Should -Be 'Units.xlsx'
        $doc.DocumentElement.SelectSingleNode("*[local-name()='ProjectCategory']").InnerText | Should -Be '1'
        $doc.DocumentElement.NamespaceURI | Should -Be 'http://schemas.microsoft.com/dynamics/2015/01/DataManagement'
    }
}

Describe 'Write-DmfPackageHeader' {
    It 'writes UTF-16 with the escaped description' {
        $p = Join-Path $tempRoot 'PackageHeader.xml'
        Write-DmfPackageHeader -Path $p -Description 'A & B <c>'
        $bytes = [System.IO.File]::ReadAllBytes($p)
        $bytes[0] | Should -Be 0xFF
        $text = [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
        $text | Should -BeLike '*<Description>A &amp; B &lt;c&gt;</Description>*'
        $text | Should -BeLike '*<PackageType>DefinitionGroup</PackageType>*'
    }
}

Describe 'Sidecar' {
    It 'round-trips template.json and returns null when absent' {
        $folder = Join-Path $tempRoot 'sidecar'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        Read-TemplateSidecar -Folder $folder | Should -BeNullOrEmpty
        Write-TemplateSidecar -Folder $folder -Sidecar ([ordered]@{ schemaVersion = 1; templateId = 'T'; lines = @(@{ entity = 'Currencies'; sysModule = 'GL' }) })
        $s = Read-TemplateSidecar -Folder $folder
        $s.templateId       | Should -Be 'T'
        $s.lines[0].entity  | Should -Be 'Currencies'
    }
}

Describe 'Get-TemplateFolders / Get-TemplateInfo' {
    BeforeAll {
        $root = Join-Path $tempRoot 'resources'
        New-Item -ItemType Directory -Path (Join-Path $root 'B template') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $root 'A package')  -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $root 'no manifest') -Force | Out-Null
        Copy-Item (Join-Path $fixtures 'minimal-utf8\Manifest.xml') (Join-Path $root 'B template\Manifest.xml')
        Copy-Item (Join-Path $fixtures 'minimal-utf8\Manifest.xml') (Join-Path $root 'A package\Manifest.xml')
        Set-Content -Path (Join-Path $root 'A package\Currencies.xlsx') -Value 'x'
        Set-Content -Path (Join-Path $root 'A package\ordering.json') -Value '{}'
    }

    It 'lists only folders with a manifest, sorted' {
        $f = @(Get-TemplateFolders -ResourcesPath $root)
        @($f | ForEach-Object Name) | Should -Be @('A package', 'B template')
    }

    It 'returns nothing for a missing root' {
        $f = @(Get-TemplateFolders -ResourcesPath (Join-Path $tempRoot 'missing'))
        $f.Count | Should -Be 0
    }

    It 'distinguishes template-only from package folders' {
        $folders = @(Get-TemplateFolders -ResourcesPath $root)
        $pkg = Get-TemplateInfo -Folder $folders[0] -Index 1
        $tpl = Get-TemplateInfo -Folder $folders[1] -Index 2
        $pkg.HasData        | Should -BeTrue
        $pkg.XlsxCount      | Should -Be 1
        $pkg.HasOrdering    | Should -BeTrue
        $tpl.HasData        | Should -BeFalse
        $tpl.EntityCount    | Should -Be 3
        $tpl.DisabledCount  | Should -Be 1
        $tpl.IsValid        | Should -BeTrue
        $tpl.ResolvedCount  | Should -BeNullOrEmpty
        $tpl.Description    | Should -Be 'Retail setup needed before store go-live'
    }
}

Describe 'Get-TemplateOrigin' {
    BeforeAll {
        $originRoot = Join-Path $tempRoot 'origin'
        New-Item -ItemType Directory -Path $originRoot -Force | Out-Null
    }

    It 'is unknown without a sidecar' {
        Get-TemplateOrigin -Folder $originRoot | Should -Be 'unknown'
    }

    It 'is custom when the sidecar says so, whatever the case or padding' {
        $f = Join-Path $originRoot 'custom'; New-Item -ItemType Directory -Path $f -Force | Out-Null
        Write-TemplateSidecar -Folder $f -Sidecar ([ordered]@{ schemaVersion = 1; templateId = 'X'; origin = ' Custom ' })
        Get-TemplateOrigin -Folder $f | Should -Be 'custom'
    }

    It 'is captured when the sidecar carries capturedFrom and no origin' {
        $f = Join-Path $originRoot 'captured'; New-Item -ItemType Directory -Path $f -Force | Out-Null
        Write-TemplateSidecar -Folder $f -Sidecar ([ordered]@{ schemaVersion = 1; templateId = 'X'; capturedFrom = 'https://x' })
        Get-TemplateOrigin -Folder $f | Should -Be 'captured'
    }

    It 'is unknown for a sidecar with neither, and honours an explicit captured' {
        $f = Join-Path $originRoot 'neither'; New-Item -ItemType Directory -Path $f -Force | Out-Null
        Write-TemplateSidecar -Folder $f -Sidecar ([ordered]@{ schemaVersion = 1; templateId = 'X' })
        Get-TemplateOrigin -Folder $f | Should -Be 'unknown'
        Get-TemplateOrigin -Folder $f -Sidecar @{ origin = 'captured' } | Should -Be 'captured'
    }

    It 'accepts a hashtable sidecar without touching disk' {
        Get-TemplateOrigin -Folder (Join-Path $originRoot 'nope') -Sidecar @{ origin = 'custom' } | Should -Be 'custom'
    }
}

Describe 'Get-TemplateInfo origin' {
    It 'exposes Origin and IsCustom from template.json' {
        $root = Join-Path $tempRoot 'resources-origin'
        New-Item -ItemType Directory -Path (Join-Path $root 'Custom one') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $root 'Plain one')  -Force | Out-Null
        Copy-Item (Join-Path $fixtures 'minimal-utf8\Manifest.xml') (Join-Path $root 'Custom one\Manifest.xml')
        Copy-Item (Join-Path $fixtures 'minimal-utf8\Manifest.xml') (Join-Path $root 'Plain one\Manifest.xml')
        Write-TemplateSidecar -Folder (Join-Path $root 'Custom one') -Sidecar ([ordered]@{ schemaVersion = 1; templateId = 'Custom one'; origin = 'custom' })

        $folders = @(Get-TemplateFolders -ResourcesPath $root)
        $custom  = Get-TemplateInfo -Folder $folders[0] -Index 1
        $plain   = Get-TemplateInfo -Folder $folders[1] -Index 2
        $custom.Origin   | Should -Be 'custom'
        $custom.IsCustom | Should -BeTrue
        $plain.Origin    | Should -Be 'unknown'
        $plain.IsCustom  | Should -BeFalse
    }
}
