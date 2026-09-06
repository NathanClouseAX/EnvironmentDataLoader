#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
    Unit tests for lib/DmfOData.ps1.  No network: Invoke-DmfRequest is mocked.
#>
Set-StrictMode -Version Latest

BeforeAll {
    $libPath = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'lib'
    . (Join-Path $libPath 'DmfOutput.ps1')
    . (Join-Path $libPath 'DmfRequest.ps1')
    . (Join-Path $libPath 'DmfOData.ps1')
    $Script:LineWidth = 80
    $Script:Version   = 'test'
}

Describe 'ConvertTo-DmfODataLiteral' {
    It 'doubles single quotes' {
        ConvertTo-DmfODataLiteral -Value "O'Brien's" | Should -Be "O''Brien''s"
    }
    It 'passes plain text through' {
        ConvertTo-DmfODataLiteral -Value '010 - System Setup' | Should -Be '010 - System Setup'
    }
    It 'accepts an empty string' {
        ConvertTo-DmfODataLiteral -Value '' | Should -Be ''
    }
}

Describe 'New-DmfODataUri' {
    It 'builds a bare collection URL' {
        New-DmfODataUri -BaseUrl 'https://x.operations.dynamics.com/' -Collection 'Currencies' |
            Should -Be 'https://x.operations.dynamics.com/data/Currencies'
    }
    It 'encodes the filter and orders cross-company first' {
        $u = New-DmfODataUri -BaseUrl 'https://x' -Collection 'CustomersV3' -CrossCompany -Filter "dataAreaId eq 'USMF'" -Top 5
        $u | Should -Be "https://x/data/CustomersV3?cross-company=true&`$filter=dataAreaId%20eq%20%27USMF%27&`$top=5"
    }
    It 'supports the Metadata service and $count' {
        New-DmfODataUri -BaseUrl 'https://x' -ServicePath 'Metadata' -Collection 'DataEntities' -Top 1 -Count |
            Should -Be 'https://x/Metadata/DataEntities?$top=1&$count=true'
    }
    It 'adds $select' {
        New-DmfODataUri -BaseUrl 'https://x' -Collection 'A' -Select 'Name,LabelId' |
            Should -Be 'https://x/data/A?$select=Name%2CLabelId'
    }
}

Describe 'Get-DmfODataValue' {
    It 'emits nothing for null' {
        $v = @(Get-DmfODataValue -Response $null)
        $v.Count | Should -Be 0
    }
    It 'emits the single row of a one-row value' {
        $v = @(Get-DmfODataValue -Response ([pscustomobject]@{ value = @([pscustomobject]@{ a = 1 }) }))
        $v.Count | Should -Be 1
        $v[0].a  | Should -Be 1
    }
    It 'emits a non-collection response as one item' {
        $v = @(Get-DmfODataValue -Response ([pscustomobject]@{ a = 1 }))
        $v.Count | Should -Be 1
    }
    It 'returns a single entity whole even when it has a scalar property called Value (Metadata Labels)' {
        # The property lookup is case-insensitive, so a scalar 'Value' must not be
        # mistaken for the 'value' collection of a collection response.
        $label = [pscustomobject]@{ '@odata.context' = 'x'; Id = '@SYS1'; Language = 'en-US'; Value = 'Currencies' }
        $v = @(Get-DmfODataValue -Response $label)
        $v.Count      | Should -Be 1
        $v[0].Value   | Should -Be 'Currencies'
        $v[0].Id      | Should -Be '@SYS1'
    }
}

Describe 'Get-DmfODataAll' {
    BeforeEach {
        Mock Write-Detail {}
        Mock Invoke-DmfRequest {
            switch -Wildcard ($Params.Uri) {
                '*page=3' { [pscustomobject]@{ value = @(@{ i = 5 }) } }
                '*page=2' { [pscustomobject]@{ value = @(@{ i = 3 }, @{ i = 4 }); '@odata.nextLink' = 'https://x/data/A?page=3' } }
                default   { [pscustomobject]@{ value = @(@{ i = 1 }, @{ i = 2 }); '@odata.nextLink' = 'https://x/data/A?page=2' } }
            }
        }
    }

    It 'follows nextLink across all pages' {
        $items = @(Get-DmfODataAll -Uri 'https://x/data/A' -Operation 'A' -Headers @{})
        $items.Count | Should -Be 5
        $items[4].i  | Should -Be 5
        Should -Invoke Invoke-DmfRequest -Times 3 -Exactly
    }

    It 'stops at MaxItems without requesting further pages' {
        $items = @(Get-DmfODataAll -Uri 'https://x/data/A' -Operation 'A' -Headers @{} -MaxItems 3)
        $items.Count | Should -Be 3
        Should -Invoke Invoke-DmfRequest -Times 2 -Exactly
    }

    It 'reports progress per page' {
        $global:DmfTestPages = @()
        [void]@(Get-DmfODataAll -Uri 'https://x/data/A' -Operation 'A' -Headers @{} -OnPage { param($p, $n) $global:DmfTestPages += "$p/$n" })
        $global:DmfTestPages | Should -Be @('1/2', '2/4', '3/5')
    }

    It 'returns nothing for an empty collection' {
        Mock Invoke-DmfRequest { [pscustomobject]@{ value = @() } }
        $items = @(Get-DmfODataAll -Uri 'https://x/data/A' -Operation 'A' -Headers @{})
        $items.Count | Should -Be 0
    }
}
