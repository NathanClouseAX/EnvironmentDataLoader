#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
    Unit tests for lib/DmfPull.ps1 (snapshot files and the run index).
#>
Set-StrictMode -Version Latest

BeforeAll {
    $repoRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $libPath  = Join-Path $repoRoot 'lib'
    . (Join-Path $libPath 'DmfOutput.ps1')
    . (Join-Path $libPath 'DmfRequest.ps1')
    . (Join-Path $libPath 'DmfAuth.ps1')
    . (Join-Path $libPath 'DmfOData.ps1')
    . (Join-Path $libPath 'DmfTemplate.ps1')
    . (Join-Path $libPath 'DmfPull.ps1')
    $Script:LineWidth = 80
    $Script:Version   = 'test'

    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("DmfPullTests_" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

    function New-TestSession {
        [pscustomobject]@{
            PSTypeName = 'Dmf.Session'; BaseUrl = 'https://mock-uat.sandbox.operations.dynamics.com'; TenantId = 't'; ClientId = 'c'
            AuthBase = 'x'; Scope = 's'; EnvironmentName = 'mock-uat'
            AccessToken = 'at'; RefreshToken = $null; ExpiresAt = (Get-Date).AddHours(1); ScopeNote = $null; RefreshCount = 0; LastRefreshAt = $null
        }
    }
    function New-Resolution {
        param([string]$Entity = 'Customer groups', [string]$Collection = 'CustomerGroups', [bool]$Company = $true, [string[]]$Keys = @('dataAreaId', 'CustomerGroupId'))
        [pscustomobject]@{ EntityName = $Entity; TargetEntity = 'CustCustomerGroupEntity'; PublicEntityName = 'CustomerGroup'; Collection = $Collection; KeyFields = $Keys; CompanySpecific = $Company; DataServiceEnabled = $true; Status = 'Resolved'; Reason = ''; Source = 'metadata' }
    }
}

AfterAll {
    if (Test-Path $tempRoot) { Remove-Item $tempRoot -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'Paths and names' {
    It 'builds data/environment/company and upper-cases the company' {
        $f = Get-DmfPullFolder -DataPath (Join-Path $tempRoot 'data') -EnvironmentName 'contoso-uat' -LegalEntityId 'usmf' -Create
        $f | Should -Be (Join-Path (Join-Path (Join-Path $tempRoot 'data') 'contoso-uat') 'USMF')
        (Test-Path $f) | Should -BeTrue
    }
    It 'names the file after the label with invalid characters removed' {
        Get-DmfSnapshotFileName -EntityName 'Country/regions' | Should -Be 'Countryregions.json'
    }
}

Describe 'Invoke-DmfEntityPull' {
    BeforeEach { Mock Write-Detail {} }

    It 'uses cross-company plus a dataAreaId filter for company-specific entities' {
        Mock Get-DmfODataAll { param($Uri, $Operation, $Headers, $MaxItems, $OnPage) @([pscustomobject]@{ dataAreaId = 'usmf'; CustomerGroupId = '10' }) }
        $r = Invoke-DmfEntityPull -Session (New-TestSession) -Resolution (New-Resolution) -LegalEntityId 'USMF'
        $r.Records.Count | Should -Be 1
        $r.Truncated     | Should -BeFalse
        $r.Uri           | Should -Be "https://mock-uat.sandbox.operations.dynamics.com/data/CustomerGroups?cross-company=true&`$filter=dataAreaId%20eq%20%27USMF%27"
    }

    It 'uses a bare collection URL for shared entities and reports truncation' {
        Mock Get-DmfODataAll { param($Uri, $Operation, $Headers, $MaxItems, $OnPage) 1..$MaxItems | ForEach-Object { [pscustomobject]@{ CurrencyCode = "C$_" } } }
        $r = Invoke-DmfEntityPull -Session (New-TestSession) -Resolution (New-Resolution -Entity 'Currencies' -Collection 'Currencies' -Company $false -Keys @('CurrencyCode')) -LegalEntityId 'USMF' -MaxRecords 3
        $r.Uri           | Should -Be 'https://mock-uat.sandbox.operations.dynamics.com/data/Currencies'
        $r.Records.Count | Should -Be 3
        $r.Truncated     | Should -BeTrue
    }
}

Describe 'Snapshot files' {
    It 'writes the envelope with sorted records and a field union, then reads it back' {
        $folder  = Join-Path $tempRoot 'snap'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        $records = @(
            [pscustomobject]@{ '@odata.etag' = 'W/"2"'; dataAreaId = 'usmf'; CustomerGroupId = '20'; Description = 'Retail' }
            [pscustomobject]@{ '@odata.etag' = 'W/"1"'; dataAreaId = 'usmf'; CustomerGroupId = '10'; Description = 'Wholesale'; Extra = 'x' }
        )
        $path = Write-DmfEntitySnapshot -Folder $folder -Resolution (New-Resolution) -Session (New-TestSession) -LegalEntityId 'usmf' -Records $records -ElapsedSeconds 1.234
        $path | Should -Be (Join-Path $folder 'Customer groups.json')
        (Test-Path "$path.tmp") | Should -BeFalse

        $s = Read-DmfEntitySnapshot -Path $path
        $s.Entity          | Should -Be 'Customer groups'
        $s.Collection      | Should -Be 'CustomerGroups'
        $s.LegalEntity     | Should -Be 'USMF'
        $s.Environment     | Should -Be 'mock-uat'
        $s.CompanySpecific | Should -BeTrue
        $s.KeyFields       | Should -Be @('dataAreaId', 'CustomerGroupId')
        $s.RecordCount     | Should -Be 2
        $s.Truncated       | Should -BeFalse
        $s.Fields          | Should -Be @('@odata.etag', 'dataAreaId', 'CustomerGroupId', 'Description', 'Extra')
        $s.Records[0].CustomerGroupId | Should -Be '10'      # sorted by key
        $s.Records[0].'@odata.etag'   | Should -Be 'W/"1"'   # verbatim
        $s.PulledAt        | Should -Match '^\d{4}-\d{2}-\d{2}T'
    }

    It 'handles an empty entity' {
        $folder = Join-Path $tempRoot 'empty'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        $path = Write-DmfEntitySnapshot -Folder $folder -Resolution (New-Resolution -Entity 'Nothing' -Collection 'Nothings') -Session (New-TestSession) -LegalEntityId 'USMF' -Records @()
        $s = Read-DmfEntitySnapshot -Path $path
        $s.RecordCount   | Should -Be 0
        $s.Records.Count | Should -Be 0
        $s.Fields.Count  | Should -Be 0
    }
}

Describe 'Run index' {
    It 'creates, merges and keeps a stale file reference on failure' {
        $folder = Join-Path $tempRoot 'index'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null

        $i = Read-DmfPullIndex -Folder $folder
        $i.entities.Count | Should -Be 0
        $i.environment = 'mock-uat'; $i.legalEntity = 'USMF'; $i.environmentUrl = 'https://x'
        Set-DmfPullIndexEntry -Index $i -EntityName 'Currencies' -Status 'Pulled' -File 'Currencies.json' -RecordCount 164 -Templates @('010 - System Setup')
        Set-DmfPullIndexEntry -Index $i -EntityName 'Workflow notes' -Status 'NotPublic' -Reason 'DataServiceEnabled=false' -Templates @('010 - System Setup')
        $i.lastRun = [ordered]@{ startedAt = 'a'; finishedAt = 'b'; templates = @('010 - System Setup'); tool = 'test' }
        Write-DmfPullIndex -Folder $folder -Index $i

        $j = Read-DmfPullIndex -Folder $folder
        $j.environment | Should -Be 'mock-uat'
        $j.entities.Count | Should -Be 2
        $j.entities['currencies']['recordCount'] | Should -Be 164
        $j.lastRun.tool | Should -Be 'test'

        # second run: one entity fails, another is new; the untouched one is retained
        Set-DmfPullIndexEntry -Index $j -EntityName 'Currencies' -Status 'Failed' -Reason 'HTTP 500' -Templates @('010 - System Setup')
        Set-DmfPullIndexEntry -Index $j -EntityName 'Units' -Status 'Pulled' -File 'Units.json' -RecordCount 5
        Write-DmfPullIndex -Folder $folder -Index $j
        $k = Read-DmfPullIndex -Folder $folder
        $k.entities.Count | Should -Be 3
        $k.entities['Currencies']['status'] | Should -Be 'Failed'
        $k.entities['Currencies']['file']   | Should -Be 'Currencies.json'
        $k.entities['Workflow notes']['status'] | Should -Be 'NotPublic'
    }
}
