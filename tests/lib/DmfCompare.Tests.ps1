#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
    Unit tests for lib/DmfCompare.ps1: value normalisation, record
    classification on in-memory records, and a folder comparison over the
    fixture snapshots in tests/fixtures/data.
#>
Set-StrictMode -Version Latest

BeforeAll {
    $repoRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $libPath  = Join-Path $repoRoot 'lib'
    . (Join-Path $libPath 'DmfOutput.ps1')
    . (Join-Path $libPath 'DmfTemplate.ps1')
    . (Join-Path $libPath 'DmfPull.ps1')
    . (Join-Path $libPath 'DmfCompare.ps1')
    $Script:LineWidth = 80
    $Script:Version   = 'test'
    $fixtures = Join-Path $repoRoot 'tests\fixtures\data'
    $envA = Join-Path $fixtures 'env-a\USMF'
    $envB = Join-Path $fixtures 'env-b\USMF'
}

Describe 'ConvertTo-DmfComparableValue' {
    It 'treats null, empty and whitespace alike' {
        ConvertTo-DmfComparableValue -Value $null | Should -BeNullOrEmpty
        ConvertTo-DmfComparableValue -Value ''    | Should -BeNullOrEmpty
        ConvertTo-DmfComparableValue -Value '   ' | Should -BeNullOrEmpty
    }
    It 'normalises booleans and Yes/No' {
        ConvertTo-DmfComparableValue -Value $true  | Should -Be 'true'
        ConvertTo-DmfComparableValue -Value 'Yes'  | Should -Be 'true'
        ConvertTo-DmfComparableValue -Value 'NO'   | Should -Be 'false'
        ConvertTo-DmfComparableValue -Value 'False'| Should -Be 'false'
    }
    It 'normalises numbers but keeps integer-looking codes verbatim' {
        ConvertTo-DmfComparableValue -Value 1.0    | Should -Be '1'
        ConvertTo-DmfComparableValue -Value '1.50' | Should -Be '1.5'
        ConvertTo-DmfComparableValue -Value 0.010  | Should -Be (ConvertTo-DmfComparableValue -Value '0.01')
        ConvertTo-DmfComparableValue -Value '0010' | Should -Be '0010'
        ConvertTo-DmfComparableValue -Value 10     | Should -Be '10'
        ConvertTo-DmfComparableValue -Value '1E3'  | Should -Be '1000'
    }
    It 'normalises ISO dates across offsets and maps the 1900 sentinel to null' {
        $a = ConvertTo-DmfComparableValue -Value '2026-01-01T00:00:00Z'
        $b = ConvertTo-DmfComparableValue -Value '2026-01-01T00:00:00.0000000+00:00'
        $c = ConvertTo-DmfComparableValue -Value ([DateTime]::new(2026, 1, 1, 0, 0, 0, [System.DateTimeKind]::Utc))
        $a | Should -Be $b
        $a | Should -Be $c
        ConvertTo-DmfComparableValue -Value '1900-01-01T00:00:00Z' | Should -BeNullOrEmpty
        ConvertTo-DmfComparableValue -Value ([DateTime]::new(1900, 1, 1, 0, 0, 0, [System.DateTimeKind]::Utc)) | Should -BeNullOrEmpty
    }
    It 'trims strings but keeps case' {
        ConvertTo-DmfComparableValue -Value ' Wholesale ' | Should -BeExactly 'Wholesale'
        ConvertTo-DmfComparableValue -Value 'wholesale'   | Should -Not -BeExactly 'Wholesale'
    }
    It 'compares raw text with -Strict' {
        ConvertTo-DmfComparableValue -Value '1.50' -Strict | Should -Be '1.50'
        ConvertTo-DmfComparableValue -Value 'Yes'  -Strict | Should -Be 'Yes'
        ConvertTo-DmfComparableValue -Value ' x '  -Strict | Should -Be ' x '
    }
    It 'serialises nested values' {
        ConvertTo-DmfComparableValue -Value @(1, 2) | Should -Be '[1,2]'
    }
}

Describe 'Compare-DmfEntityRecords' {
    BeforeAll {
        $ref = @(
            [pscustomobject]@{ '@odata.etag' = 'a'; Code = 'A'; Name = 'Alpha'; Qty = 1;   ModifiedDateTime = '2026-01-01T00:00:00Z' }
            [pscustomobject]@{ '@odata.etag' = 'b'; Code = 'B'; Name = 'Beta';  Qty = 2.5; ModifiedDateTime = '2026-01-01T00:00:00Z' }
            [pscustomobject]@{ '@odata.etag' = 'c'; Code = 'C'; Name = 'Gone';  Qty = 3;   ModifiedDateTime = '2026-01-01T00:00:00Z' }
        )
        $dif = @(
            [pscustomobject]@{ '@odata.etag' = 'x'; Code = 'A'; Name = 'Alpha!'; Qty = '1.0'; ModifiedDateTime = '2026-06-01T00:00:00Z'; Extra = 'new' }
            [pscustomobject]@{ '@odata.etag' = 'y'; Code = 'B'; Name = 'Beta';   Qty = '2.50'; ModifiedDateTime = '2026-06-01T00:00:00Z'; Extra = 'new' }
            [pscustomobject]@{ '@odata.etag' = 'z'; Code = 'D'; Name = 'Delta';  Qty = 4;     ModifiedDateTime = '2026-06-01T00:00:00Z'; Extra = 'new' }
        )
    }

    It 'classifies added, removed, changed and unchanged with default ignores' {
        $r = Compare-DmfEntityRecords -Reference $ref -Difference $dif -KeyFields @('Code')
        $r.Keyless               | Should -BeFalse
        $r.Added.Count           | Should -Be 1
        $r.Added[0].Code         | Should -Be 'D'
        $r.Removed.Count         | Should -Be 1
        $r.Removed[0].Code       | Should -Be 'C'
        $r.Changed.Count         | Should -Be 1
        $r.Changed[0].KeyValues.Code | Should -Be 'A'
        $r.Changed[0].Fields.Count   | Should -Be 1
        $r.Changed[0].Fields[0].Field           | Should -Be 'Name'
        $r.Changed[0].Fields[0].ReferenceValue  | Should -Be 'Alpha'
        $r.Changed[0].Fields[0].DifferenceValue | Should -Be 'Alpha!'
        $r.UnchangedCount        | Should -Be 1
        $r.SchemaDrift.DifferenceOnly | Should -Be @('Extra')
        $r.SchemaDrift.ReferenceOnly.Count | Should -Be 0
        $r.IgnoredFields | Should -Contain '@odata.etag'
        $r.IgnoredFields | Should -Contain 'ModifiedDateTime'
    }

    It 'reports every field with -Strict and no ignores' {
        $r = Compare-DmfEntityRecords -Reference $ref -Difference $dif -KeyFields @('Code') -IgnorePatterns @() -Strict
        $a = $r.Changed | Where-Object { $_.KeyValues.Code -eq 'A' }
        @($a.Fields | ForEach-Object Field) | Should -Contain 'Qty'
        @($a.Fields | ForEach-Object Field) | Should -Contain '@odata.etag'
        @($a.Fields | ForEach-Object Field) | Should -Contain 'ModifiedDateTime'
        $b = $r.Changed | Where-Object { $_.KeyValues.Code -eq 'B' }
        @($b.Fields | ForEach-Object Field) | Should -Contain 'Qty'      # 2.5 vs '2.50' differ only in strict mode
    }

    It 'falls back to a keyless comparison and finds duplicate keys' {
        $r = Compare-DmfEntityRecords -Reference $ref -Difference $dif -KeyFields @() -IgnorePatterns @('@odata.etag', 'ModifiedDateTime', 'Extra')
        $r.Keyless          | Should -BeTrue
        $r.KeyFields        | Should -Be @('Code', 'Name', 'Qty')
        $r.Changed.Count    | Should -Be 0
        $r.Removed.Count    | Should -Be 2        # A (name changed) and C
        $r.Added.Count      | Should -Be 2        # A' and D
        $r.UnchangedCount   | Should -Be 1        # B

        $dup = Compare-DmfEntityRecords -Reference @($ref + $ref[0]) -Difference $dif -KeyFields @('Code')
        $dup.KeyCollisions.Reference.Count | Should -Be 1
        $dup.ReferenceCount | Should -Be 4
    }

    It 'drops an ignored field from the key' {
        $a = @([pscustomobject]@{ dataAreaId = 'usmf'; Id = '10'; Description = 'x' })
        $b = @([pscustomobject]@{ dataAreaId = 'dat';  Id = '10'; Description = 'y' })
        $r = Compare-DmfEntityRecords -Reference $a -Difference $b -KeyFields @('dataAreaId', 'Id') -IgnorePatterns @('dataAreaId')
        $r.KeyFields       | Should -Be @('Id')
        $r.Changed.Count   | Should -Be 1
        $r.Changed[0].Fields[0].Field | Should -Be 'Description'
    }

    It 'handles empty sides' {
        $r = Compare-DmfEntityRecords -Reference @() -Difference $dif -KeyFields @('Code')
        $r.Added.Count | Should -Be 3
        $r.Removed.Count | Should -Be 0
        $r2 = Compare-DmfEntityRecords -Reference @() -Difference @() -KeyFields @('Code')
        $r2.UnchangedCount | Should -Be 0
    }
}

Describe 'Compare-DmfSnapshotFolders (fixtures)' {
    BeforeEach { Mock Write-Warn {} }

    It 'compares the two fixture folders entity by entity' {
        $c = Compare-DmfSnapshotFolders -ReferencePath $envA -DifferencePath $envB
        $c.Entities.Count | Should -Be 4
        $by = @{}; foreach ($e in $c.Entities) { $by[$e.Entity] = $e }

        $cur = $by['Currencies']
        $cur.Status               | Should -Be 'Compared'
        $cur.Result.Added.Count   | Should -Be 1     # USD
        $cur.Result.Removed.Count | Should -Be 1     # OLD
        $cur.Result.Changed.Count | Should -Be 1     # AED name; EUR equal after normalisation (0.010, true/Yes, trimmed)
        $cur.Result.Changed[0].Fields.Count | Should -Be 1
        $cur.Result.Changed[0].Fields[0].Field | Should -Be 'Name'
        $cur.Result.SchemaDrift.DifferenceOnly | Should -Be @('Symbol')

        $cg = $by['Customer groups']
        $cg.Result.Changed.Count   | Should -Be 0   # 'usmf' vs 'USMF' key, ' Wholesale ' vs 'Wholesale'
        $cg.Result.UnchangedCount  | Should -Be 2

        $by['Units'].Status  | Should -Be 'OnlyInReference'
        $by['Units'].Reason  | Should -BeLike 'Failed*HTTP 500*'
        $by['Sites'].Status  | Should -Be 'OnlyInDifference'
        $by['Sites'].Reason  | Should -BeLike 'NotPublic*'
    }

    It 'filters entities by wildcard and honours -KeyOverride' {
        $c = Compare-DmfSnapshotFolders -ReferencePath $envA -DifferencePath $envB -Entity 'Cust*'
        $c.Entities.Count | Should -Be 1
        $c.Entities[0].Entity | Should -Be 'Customer groups'

        $k = Compare-DmfSnapshotFolders -ReferencePath $envA -DifferencePath $envB -Entity 'Currencies' -KeyOverride @{ 'Currencies' = @('Name') }
        $k.Entities[0].Result.KeyFields | Should -Be @('Name')
        $k.Entities[0].Result.Added.Count | Should -Be 2   # renamed AED + USD
    }

    It 'filters by template (package) using either side''s pull index' {
        $c = Compare-DmfSnapshotFolders -ReferencePath $envA -DifferencePath $envB -Template '010*'
        @($c.Entities | ForEach-Object Entity) | Should -Be @('Currencies', 'Customer groups')
        $c.TemplateFilter | Should -Be @('010*')

        $g = Compare-DmfSnapshotFolders -ReferencePath $envA -DifferencePath $envB -Template '020 - GL Shared', '300 - Inventory'
        @($g.Entities | ForEach-Object Entity) | Should -Be @('Currencies', 'Sites', 'Units')
    }

    It 'rolls results up per template' {
        $c = Compare-DmfSnapshotFolders -ReferencePath $envA -DifferencePath $envB
        $by = @{}; foreach ($t in $c.Templates) { $by[$t.Template] = $t }
        @($c.Templates | ForEach-Object Template) | Should -Be @('010 - System Setup', '020 - GL Shared', '300 - Inventory')

        $t010 = $by['010 - System Setup']
        $t010.Entities     | Should -Be @('Currencies', 'Customer groups')
        $t010.Compared     | Should -Be 2
        $t010.Identical    | Should -Be 1
        $t010.WithChanges  | Should -Be 1
        $t010.Added        | Should -Be 1
        $t010.Removed      | Should -Be 1
        $t010.Changed      | Should -Be 1
        $t010.InReferenceIndex  | Should -BeTrue
        $t010.InDifferenceIndex | Should -BeTrue

        $t020 = $by['020 - GL Shared']
        $t020.Entities        | Should -Be @('Currencies', 'Units')
        $t020.OnlyInReference | Should -Be 1        # Units failed on the Difference side
        $t020.WithChanges     | Should -Be 1        # Currencies is shared with 010

        $t300 = $by['300 - Inventory']
        $t300.OnlyInDifference | Should -Be 1       # Sites not public on the Reference side
        $t300.Compared         | Should -Be 0
    }

    It 'is empty when a folder is compared with itself' {
        $c = Compare-DmfSnapshotFolders -ReferencePath $envA -DifferencePath $envA
        foreach ($e in $c.Entities) {
            $e.Status | Should -Be 'Compared'
            $e.Result.Added.Count + $e.Result.Removed.Count + $e.Result.Changed.Count | Should -Be 0
        }
    }

    It 'flattens findings for CSV' {
        $c = Compare-DmfSnapshotFolders -ReferencePath $envA -DifferencePath $envB
        $f = @(ConvertTo-DmfCompareFindings -Comparison $c)
        @($f | Where-Object ChangeType -eq 'Changed').Count               | Should -Be 1
        @($f | Where-Object ChangeType -eq 'Added').Count                 | Should -Be 1
        @($f | Where-Object ChangeType -eq 'Removed').Count               | Should -Be 1
        @($f | Where-Object ChangeType -eq 'SchemaDrift').Count           | Should -Be 1
        @($f | Where-Object ChangeType -eq 'EntityOnlyInReference').Count | Should -Be 1
        @($f | Where-Object ChangeType -eq 'EntityOnlyInDifference').Count| Should -Be 1
        ($f | Where-Object ChangeType -eq 'Changed').Key   | Should -Be 'CurrencyCode=AED'
        ($f | Where-Object ChangeType -eq 'Removed').Record.CurrencyCode | Should -Be 'OLD'
    }
}
