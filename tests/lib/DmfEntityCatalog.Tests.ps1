#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
    Unit tests for lib/DmfEntityCatalog.ps1.  A small catalog document is built
    in the temp folder; the shipped resources/entity-catalog.json is also read
    to confirm the real file matches the shape the reader expects.
#>
Set-StrictMode -Version Latest

BeforeAll {
    $repoRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $libPath  = Join-Path $repoRoot 'lib'
    . (Join-Path $libPath 'DmfOutput.ps1')
    . (Join-Path $libPath 'DmfEntityCatalog.ps1')
    $Script:LineWidth = 80
    $Script:Version   = 'test'

    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("DmfCatalogTests_" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

    $fixture = [ordered]@{
        schemaVersion = 1
        sourceVersion = '10.0.0.1'
        generatedAt   = '2026-01-01T00:00:00Z'
        entityCount   = 4
        entities      = [ordered]@{
            WidgetEntity = [ordered]@{
                label = 'Widgets'; isPublic = $true; dmEnabled = $true; category = 'Reference'
                publicEntityName = 'Widget'; collection = 'Widgets'; companySpecific = $true
                rootTable = 'WidgetTable'; keyFields = @('dataAreaId', 'WidgetId')
            }
            GadgetEntity = [ordered]@{
                label = 'Gadgets'; isPublic = $false; dmEnabled = $true; category = 'Reference'
                publicEntityName = ''; collection = ''; companySpecific = $false
                rootTable = 'GadgetTable'; keyFields = @('GadgetId')
            }
            # Same label as GadgetEntity but not DMF-enabled: the DMF one wins.
            GadgetBiEntity = [ordered]@{
                label = 'Gadgets'; isPublic = $true; dmEnabled = $false; category = ''
                publicEntityName = 'GadgetBi'; collection = 'GadgetBis'; companySpecific = $false
                rootTable = 'GadgetTable'; keyFields = @('GadgetId')
            }
            SparseEntity = [ordered]@{ label = 'Sparse thing' }
        }
    }
    $catalogPath = Join-Path $tempRoot 'entity-catalog.json'
    [System.IO.File]::WriteAllText($catalogPath, ($fixture | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($false)))

    $shippedCatalogPath = Join-Path $repoRoot 'resources\entity-catalog.json'
}

AfterAll {
    if (Test-Path $tempRoot) { Remove-Item $tempRoot -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'Get-DmfEntityCatalog' {
    It 'reads a catalog document' {
        $doc = Get-DmfEntityCatalog -Path $catalogPath
        $doc.sourceVersion | Should -BeExactly '10.0.0.1'
        $doc.entityCount   | Should -Be 4
    }
    It 'returns null for a missing file' {
        Get-DmfEntityCatalog -Path (Join-Path $tempRoot 'nope.json') | Should -BeNullOrEmpty
    }
    It 'warns and returns null for an unparseable file' {
        $bad = Join-Path $tempRoot 'bad.json'
        Set-Content -LiteralPath $bad -Value '{ not json' -Encoding UTF8
        Mock Write-Warn {}
        Get-DmfEntityCatalog -Path $bad | Should -BeNullOrEmpty
    }
}

Describe 'ConvertTo-DmfCatalogEntities' {
    It 'flattens the document and carries the AOT name' {
        $entities = @(ConvertTo-DmfCatalogEntities -Document (Get-DmfEntityCatalog -Path $catalogPath))
        $entities.Count | Should -Be 4
        # Select-Object -First 1, not @(...): member access on a one-element
        # array wraps the value, which hides an empty KeyFields array.
        $widget = $entities | Where-Object { $_.Name -eq 'WidgetEntity' } | Select-Object -First 1
        $widget.Label           | Should -BeExactly 'Widgets'
        $widget.IsPublic        | Should -BeTrue
        $widget.CompanySpecific | Should -BeTrue
        $widget.Collection      | Should -BeExactly 'Widgets'
        @($widget.KeyFields)    | Should -Be @('dataAreaId', 'WidgetId')
    }
    It 'defaults every missing property rather than throwing under strict mode' {
        $entities = @(ConvertTo-DmfCatalogEntities -Document (Get-DmfEntityCatalog -Path $catalogPath))
        $sparse = $entities | Where-Object { $_.Name -eq 'SparseEntity' } | Select-Object -First 1
        $sparse.IsPublic   | Should -BeFalse
        $sparse.Collection | Should -BeExactly ''
        @($sparse.KeyFields).Count | Should -Be 0
    }
    It 'returns nothing for a null document' {
        @(ConvertTo-DmfCatalogEntities -Document $null).Count | Should -Be 0
    }
}

Describe 'Get-DmfCatalogNormalizedLabel' {
    It 'drops version tokens and singularises' {
        Get-DmfCatalogNormalizedLabel -Text 'Postal codes V3' | Should -BeExactly 'code postal'
    }
    It 'is insensitive to word order and punctuation' {
        (Get-DmfCatalogNormalizedLabel -Text 'Country/regions') |
            Should -BeExactly (Get-DmfCatalogNormalizedLabel -Text 'regions country')
    }
    It 'keeps a trailing double-s intact' {
        Get-DmfCatalogNormalizedLabel -Text 'Address' | Should -BeExactly 'address'
    }
    It 'returns empty for blank input' {
        Get-DmfCatalogNormalizedLabel -Text '' | Should -BeExactly ''
    }
}

Describe 'Find-DmfCatalogEntity' {
    BeforeAll {
        $entities = @(ConvertTo-DmfCatalogEntities -Document (Get-DmfEntityCatalog -Path $catalogPath))
        $lookup   = New-DmfCatalogLabelLookup -Entities $entities
    }
    It 'finds an entity by its exact label' {
        (Find-DmfCatalogEntity -Label 'Widgets' -Index $lookup).Name | Should -BeExactly 'WidgetEntity'
    }
    It 'falls back to the normalised label' {
        (Find-DmfCatalogEntity -Label 'widget' -Index $lookup).Name | Should -BeExactly 'WidgetEntity'
    }
    It 'prefers the DMF-enabled entity when a label is shared' {
        (Find-DmfCatalogEntity -Label 'Gadgets' -Index $lookup).Name | Should -BeExactly 'GadgetEntity'
    }
    It 'returns null for an unknown or blank label' {
        Find-DmfCatalogEntity -Label 'No such entity' -Index $lookup | Should -BeNullOrEmpty
        Find-DmfCatalogEntity -Label ''               -Index $lookup | Should -BeNullOrEmpty
    }
    It 'builds its own index when given entities instead' {
        (Find-DmfCatalogEntity -Label 'Widgets' -Entities $entities).Name | Should -BeExactly 'WidgetEntity'
    }
}

Describe 'The shipped catalog' {
    It 'is present and readable' {
        $doc = Get-DmfEntityCatalog -Path $shippedCatalogPath
        $doc               | Should -Not -BeNullOrEmpty
        $doc.sourceVersion | Should -Not -BeNullOrEmpty
    }
    It 'resolves a well-known entity with the fields the OData path needs' {
        $entities = @(ConvertTo-DmfCatalogEntities -Document (Get-DmfEntityCatalog -Path $shippedCatalogPath))
        $entities.Count | Should -BeGreaterThan 1000

        $currencies = Find-DmfCatalogEntity -Label 'Currencies' -Entities $entities
        $currencies.Name       | Should -BeExactly 'CurrencyEntity'
        $currencies.IsPublic   | Should -BeTrue
        $currencies.Collection | Should -BeExactly 'Currencies'
        @($currencies.KeyFields) | Should -Contain 'CurrencyCode'
    }
}
