#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
    Unit tests for lib/DmfMetadata.ps1.  The Metadata service is mocked at
    Get-DmfODataAll, so the resolution logic, cache and harvest run for real.
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
    . (Join-Path $libPath 'DmfMetadata.ps1')
    $Script:LineWidth = 80
    $Script:Version   = 'test'

    $fixtures = Join-Path $repoRoot 'tests\fixtures'
    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("DmfMetadataTests_" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

    function New-TestSession {
        [pscustomobject]@{
            PSTypeName = 'Dmf.Session'; BaseUrl = 'https://mock.operations.dynamics.com'; TenantId = 't'; ClientId = 'c'
            AuthBase = 'https://login.microsoftonline.com/t/oauth2/v2.0'; Scope = 's'; EnvironmentName = 'mock'
            AccessToken = 'at'; RefreshToken = $null; ExpiresAt = (Get-Date).AddHours(1); ScopeNote = $null; RefreshCount = 0; LastRefreshAt = $null
        }
    }

    # Canned Metadata service.  $global:MetaMock.ValueFilter = $false makes the
    # Labels Value filter throw, exercising the label-index fallback.
    $global:MetaMock = @{ ValueFilter = $true; Calls = [System.Collections.Generic.List[string]]::new() }
    $global:MetaEntities = @(
        [pscustomobject]@{ Name = 'CurrencyEntity';          PublicEntityName = 'Currency';      PublicCollectionName = 'Currencies';     LabelId = '@SYS1'; DataServiceEnabled = $true;  DataManagementEnabled = $true;  EntityCategory = 'Reference' }
        [pscustomobject]@{ Name = 'CustCustomerGroupEntity'; PublicEntityName = 'CustomerGroup'; PublicCollectionName = 'CustomerGroups'; LabelId = '@SYS2'; DataServiceEnabled = $true;  DataManagementEnabled = $true;  EntityCategory = 'Reference' }
        [pscustomobject]@{ Name = 'WorkflowNoteEntity';      PublicEntityName = '';              PublicCollectionName = '';               LabelId = '@SYS3'; DataServiceEnabled = $false; DataManagementEnabled = $true;  EntityCategory = 'Reference' }
        [pscustomobject]@{ Name = 'UnitEntityV1';            PublicEntityName = 'UnitV1';        PublicCollectionName = 'UnitsV1';        LabelId = '@SYS4'; DataServiceEnabled = $true;  DataManagementEnabled = $true;  EntityCategory = 'Reference' }
        [pscustomobject]@{ Name = 'UnitEntityV2';            PublicEntityName = 'UnitV2';        PublicCollectionName = 'UnitsV2';        LabelId = '@SYS4'; DataServiceEnabled = $true;  DataManagementEnabled = $true;  EntityCategory = 'Reference' }
        [pscustomobject]@{ Name = 'InternalOnlyEntity';      PublicEntityName = 'InternalOnly';  PublicCollectionName = 'InternalOnlys';  LabelId = '@SYS5'; DataServiceEnabled = $true;  DataManagementEnabled = $false; EntityCategory = 'Master' }
    )
    $global:MetaLabels = @{ '@SYS1' = 'Currencies'; '@SYS2' = 'Customer groups'; '@SYS3' = 'Workflow notes'; '@SYS4' = 'Units'; '@SYS5' = 'Internal only' }
}

AfterAll {
    if (Test-Path $tempRoot) { Remove-Item $tempRoot -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'Entity map cache' {
    It 'starts empty when the file does not exist' {
        $m = Get-DmfEntityMap -Path (Join-Path $tempRoot 'none\entity-map.json')
        $m.Entities.Count | Should -Be 0
        Test-DmfEntityMapResolved -Map $m -EntityName 'Currencies' | Should -BeFalse
    }

    It 'saves, reloads, and treats a not-public entry as resolved' {
        $path = Join-Path $tempRoot 'map1\entity-map.json'
        $m = Get-DmfEntityMap -Path $path
        [void](Set-DmfEntityMapEntry -Map $m -EntityName 'Currencies' -Values @{ targetEntity = 'CurrencyEntity'; publicCollectionName = 'Currencies'; dataServiceEnabled = $true; keyFields = @('CurrencyCode'); source = 'metadata' })
        [void](Set-DmfEntityMapEntry -Map $m -EntityName 'Workflow notes' -Values @{ targetEntity = 'WorkflowNoteEntity'; dataServiceEnabled = $false; source = 'metadata' })
        Save-DmfEntityMap -Map $m | Should -Be $path
        (Test-Path $path) | Should -BeTrue
        (Test-Path "$path.tmp") | Should -BeFalse

        $back = Get-DmfEntityMap -Path $path
        $back.Entities.Count | Should -Be 2
        $back.Entities['currencies'].keyFields | Should -Be @('CurrencyCode')
        $back.Entities['Currencies'].resolvedAt | Should -Not -BeNullOrEmpty
        Test-DmfEntityMapResolved -Map $back -EntityName 'Currencies'     | Should -BeTrue
        Test-DmfEntityMapResolved -Map $back -EntityName 'Workflow notes' | Should -BeTrue
        Test-DmfEntityMapResolved -Map $back -EntityName 'Missing'        | Should -BeFalse
    }

    It 'keeps a manual entry on disk when a metadata resolution tries to overwrite it' {
        $path = Join-Path $tempRoot 'map2\entity-map.json'
        $disk = Get-DmfEntityMap -Path $path
        [void](Set-DmfEntityMapEntry -Map $disk -EntityName 'Units' -Values @{ targetEntity = 'UnitEntityV2'; publicCollectionName = 'UnitsV2'; dataServiceEnabled = $true; source = 'manual' })
        [void](Save-DmfEntityMap -Map $disk)

        $mem = Get-DmfEntityMap -Path (Join-Path $tempRoot 'elsewhere.json')
        [void](Set-DmfEntityMapEntry -Map $mem -EntityName 'Units' -Values @{ targetEntity = 'UnitEntityV1'; publicCollectionName = 'UnitsV1'; dataServiceEnabled = $true; source = 'metadata' })
        [void](Save-DmfEntityMap -Map $mem -Path $path)

        (Get-DmfEntityMap -Path $path).Entities['Units'].targetEntity | Should -Be 'UnitEntityV2'
    }

    It 'only fills empty fields of a manual entry unless -Force' {
        $m = Get-DmfEntityMap -Path (Join-Path $tempRoot 'x.json')
        [void](Set-DmfEntityMapEntry -Map $m -EntityName 'Units' -Values @{ targetEntity = 'UnitEntityV2'; source = 'manual' })
        [void](Set-DmfEntityMapEntry -Map $m -EntityName 'Units' -Values @{ targetEntity = 'UnitEntityV1'; publicCollectionName = 'UnitsV1'; source = 'metadata' })
        $m.Entities['Units'].targetEntity         | Should -Be 'UnitEntityV2'
        $m.Entities['Units'].publicCollectionName | Should -Be 'UnitsV1'
        $m.Entities['Units'].source               | Should -Be 'manual'
        [void](Set-DmfEntityMapEntry -Map $m -EntityName 'Units' -Values @{ targetEntity = 'UnitEntityV1' } -Force)
        $m.Entities['Units'].targetEntity         | Should -Be 'UnitEntityV1'
    }
}

Describe 'Update-DmfEntityMapFromManifests' {
    It 'harvests TargetEntity pairs from manifests on disk and is idempotent' {
        # Use the one committed package folder, not all of resources/: captured
        # template folders come and go and would change the counts.
        $package = Join-Path $repoRoot 'resources\010 - System Setup'
        $m = Get-DmfEntityMap -Path (Join-Path $tempRoot 'h.json')
        $r = Update-DmfEntityMapFromManifests -Path $package, (Join-Path $fixtures 'manifests') -Map $m
        $r.Files  | Should -Be 3
        $r.Added  | Should -Be 92                        # the fixture's 'Exchange rates' is already in the real manifest
        $r.AlreadyKnown | Should -BeGreaterOrEqual 1
        $m.Entities['Currencies'].targetEntity    | Should -Be 'CurrencyEntity'
        $m.Entities['Currencies'].source          | Should -Be 'manifest'
        $m.Entities['Exchange rates'].targetEntity | Should -Be 'ExchangeRateEntity'
        Test-DmfEntityMapResolved -Map $m -EntityName 'Currencies' | Should -BeFalse   # target only, not yet OData-complete

        $again = Update-DmfEntityMapFromManifests -Path $package -Map $m
        $again.Added        | Should -Be 0
        $again.AlreadyKnown | Should -Be 92
    }
}

Describe 'Name tokens and candidates' {
    It 'tokenises labels and AOT names comparably' {
        Get-DmfNameTokens -Text 'Customer groups'          | Should -Be @('customer', 'group')
        Get-DmfNameTokens -Text 'CustCustomerGroupEntity'  | Should -Be @('cust', 'customer', 'group')
        Get-DmfNameTokens -Text 'Currencies'               | Should -Be @('currency')
        Get-DmfNameTokens -Text 'Country/regions'          | Should -Be @('country', 'region')
        Get-DmfNameTokens -Text 'Postal codes V3'          | Should -Be @('postal', 'code')
        Get-DmfNameTokens -Text 'Address and contact information purpose' | Should -Be @('address', 'contact', 'information', 'purpose')
    }
    It 'ranks the obvious entity first and drops unrelated ones' {
        $c = @(Find-DmfEntityCandidatesByName -Label 'Customer groups' -Entities $global:MetaEntities)
        $c.Count   | Should -BeGreaterOrEqual 1
        $c[0].Name | Should -Be 'CustCustomerGroupEntity'
        @($c | ForEach-Object Name) | Should -Not -Contain 'CurrencyEntity'
        @(Find-DmfEntityCandidatesByName -Label 'Zebra stripes' -Entities $global:MetaEntities).Count | Should -Be 0
    }
}

Describe 'Metadata label reads' {
    BeforeEach {
        Mock Write-Detail {}
        $Script:DmfMetadataLabelFilterOk = $null
        $Script:DmfMetadataLabelTextCache = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
        $global:MetaMock.Calls.Clear()
        Mock Get-DmfODataAll {
            $u = [System.Uri]::UnescapeDataString($Uri)
            $global:MetaMock.Calls.Add($u)
            if ($u -match "Labels\(Id='([^']+)',Language='([^']+)'\)") {
                $id = $Matches[1]
                if ($global:MetaLabels.ContainsKey($id)) { return [pscustomobject]@{ Id = $id; Language = $Matches[2]; Value = $global:MetaLabels[$id] } }
                throw "[Metadata Label $id] HTTP 404: not found"
            }
            throw "[Metadata Labels batch 1] HTTP 501: The label id and language must be provided as part of the key segment."
        }
    }
    It 'switches to key segments after one refused filter batch and caches results' {
        $t = Get-DmfMetadataLabels -Session (New-TestSession) -LabelIds @('@SYS1', '@SYS2', '@NOPE') -Language 'en-US'
        $t['@SYS1'] | Should -Be 'Currencies'
        $t['@SYS2'] | Should -Be 'Customer groups'
        $t.ContainsKey('@NOPE') | Should -BeFalse
        $Script:DmfMetadataLabelFilterOk | Should -BeFalse
        @($global:MetaMock.Calls | Where-Object { $_ -like '*$filter*' }).Count | Should -Be 1
        @($global:MetaMock.Calls | Where-Object { $_ -like '*Labels(Id=*' }).Count | Should -Be 3
        # second call: no filter attempt, no repeat reads
        $global:MetaMock.Calls.Clear()
        $t2 = Get-DmfMetadataLabels -Session (New-TestSession) -LabelIds @('@SYS1', '@SYS3') -Language 'en-US'
        $t2['@SYS3'] | Should -Be 'Workflow notes'
        @($global:MetaMock.Calls | Where-Object { $_ -like '*$filter*' }).Count | Should -Be 0
        @($global:MetaMock.Calls).Count | Should -Be 1
    }
    It 'encodes the key segment' {
        [void](Get-DmfMetadataLabel -Session (New-TestSession) -LabelId '@FieldServiceIntegrationLabels:Foo' -Language 'en-US')
        $global:MetaMock.Calls[0] | Should -Be "https://mock.operations.dynamics.com/Metadata/Labels(Id='@FieldServiceIntegrationLabels:Foo',Language='en-US')"
    }
}

Describe 'Resolve-DmfEntity' {
    BeforeEach {
        Mock Write-Warn   {}
        Mock Write-Info   {}
        Mock Write-Detail {}
        $global:MetaMock.ValueFilter = $true
        $global:MetaMock.Calls.Clear()
        $Script:DmfMetadataEntityCache = @{}

        $Script:DmfMetadataLabelFilterOk = $null
        $Script:DmfMetadataLabelTextCache = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
        $Script:DmfMetadataPublicCache = @{}

        Mock Get-DmfODataAll {
            $u = [System.Uri]::UnescapeDataString($Uri)
            $global:MetaMock.Calls.Add($u)
            if ($u -like '*/Metadata/DataEntities*') {
                $rows = $global:MetaEntities
                if ($u -match "DataEntities\('([^']+)'\)") { $n = $Matches[1]; $rows = @($rows | Where-Object Name -eq $n); if ($rows.Count -eq 0) { throw "[x] HTTP 404: not found" }; return $rows[0] }
                if ($u -match "Name eq '([^']+)'") { $n = $Matches[1]; $rows = @($rows | Where-Object Name -eq $n) }
                if ($u -match "LabelId eq") {
                    $ids = [regex]::Matches($u, "LabelId eq '([^']+)'") | ForEach-Object { $_.Groups[1].Value }
                    $rows = @($rows | Where-Object { $ids -contains $_.LabelId })
                }
                return $rows
            }
            if ($u -like '*/Metadata/Labels*') {
                # Real service: $filter is refused with 501 unless $global:MetaMock.ValueFilter
                # is $true (a hypothetical version that supports it); key segments always work.
                if ($u -match "Labels\(Id='([^']+)',Language='([^']+)'\)") {
                    $id = $Matches[1]
                    if ($global:MetaLabels.ContainsKey($id)) { return [pscustomobject]@{ Id = $id; Language = $Matches[2]; Value = $global:MetaLabels[$id] } }
                    throw "[Metadata Label $id] HTTP 404: label not found"
                }
                if (-not $global:MetaMock.ValueFilter) { throw "[Metadata Labels] HTTP 501: The label id and language must be provided as part of the key segment." }
                if ($u -match "Value eq '([^']+)'") {
                    $v = $Matches[1]
                    return @($global:MetaLabels.GetEnumerator() | Where-Object { $_.Value -eq $v } | ForEach-Object { [pscustomobject]@{ Id = $_.Key; Language = 'en-US'; Value = $_.Value } })
                }
                $ids = [regex]::Matches($u, "Id eq '([^']+)'") | ForEach-Object { $_.Groups[1].Value }
                return @($ids | Where-Object { $global:MetaLabels.ContainsKey($_) } | ForEach-Object { [pscustomobject]@{ Id = $_; Language = 'en-US'; Value = $global:MetaLabels[$_] } })
            }
            if ($u -like '*/Metadata/PublicEntities*') {
                if ($u -match "Name eq '([^']+)'") {
                    $n = $Matches[1]
                    $props = @([pscustomobject]@{ Name = 'CurrencyCode'; IsKey = $true }, [pscustomobject]@{ Name = 'Name'; IsKey = $false })
                    if ($n -eq 'CustomerGroup') { $props = @([pscustomobject]@{ Name = 'dataAreaId'; IsKey = $true }, [pscustomobject]@{ Name = 'CustomerGroupId'; IsKey = $true }, [pscustomobject]@{ Name = 'Description'; IsKey = $false }) }
                    return @([pscustomobject]@{ Name = $n; EntitySetName = "${n}s"; Properties = $props })
                }
            }
            throw "unhandled $u"
        }
    }

    It 'resolves by TargetEntity, by label, marks not-public, and fills the map' {
        $m = Get-DmfEntityMap -Path (Join-Path $tempRoot 'r1.json')
        $lines = @(
            [pscustomobject]@{ EntityName = 'Currencies';      TargetEntity = 'CurrencyEntity' }
            [pscustomobject]@{ EntityName = 'Customer groups'; TargetEntity = $null }
            [pscustomobject]@{ EntityName = 'Workflow notes';  TargetEntity = $null }
            [pscustomobject]@{ EntityName = 'Nowhere';         TargetEntity = $null }
        )
        $r = @(Resolve-DmfEntity -Lines $lines -Session (New-TestSession) -Map $m)
        $r.Count | Should -Be 4
        $r[0].Status | Should -Be 'Resolved';  $r[0].Collection | Should -Be 'Currencies';     $r[0].KeyFields | Should -Be @('CurrencyCode'); $r[0].CompanySpecific | Should -BeFalse; $r[0].Source | Should -Be 'metadata'
        $r[1].Status | Should -Be 'Resolved';  $r[1].Collection | Should -Be 'CustomerGroups'; $r[1].TargetEntity | Should -Be 'CustCustomerGroupEntity'; $r[1].CompanySpecific | Should -BeTrue; $r[1].KeyFields | Should -Be @('dataAreaId', 'CustomerGroupId')
        $r[2].Status | Should -Be 'NotPublic'; $r[2].TargetEntity | Should -Be 'WorkflowNoteEntity'
        $r[3].Status | Should -Be 'Unresolved'; $r[3].Reason | Should -BeLike '*no data entity with this label*'
        $m.Entities.Count | Should -Be 3
        $m.Entities['Customer groups'].companySpecific | Should -BeTrue
        $m.Entities['Currencies'].resolvedFrom          | Should -Be 'mock'
    }

    It 'answers from the cache on a second run without touching the network' {
        $m = Get-DmfEntityMap -Path (Join-Path $tempRoot 'r2.json')
        $lines = @([pscustomobject]@{ EntityName = 'Currencies'; TargetEntity = $null })
        [void]@(Resolve-DmfEntity -Lines $lines -Session (New-TestSession) -Map $m)
        $global:MetaMock.Calls.Clear()
        $r = @(Resolve-DmfEntity -Lines $lines -Session (New-TestSession) -Map $m)
        $r[0].Status | Should -Be 'Resolved'
        $r[0].Source | Should -Be 'cache'
        $global:MetaMock.Calls.Count | Should -Be 0
        # -Offline with an empty map reports Unresolved without throwing
        $o = @(Resolve-DmfEntity -Lines @([pscustomobject]@{ EntityName = 'Units'; TargetEntity = $null }) -Map $m -Offline)
        $o[0].Status | Should -Be 'Unresolved'
        $o[0].Reason | Should -BeLike '*offline*'
    }

    It 'flags an ambiguous label' {
        $m = Get-DmfEntityMap -Path (Join-Path $tempRoot 'r3.json')
        $r = @(Resolve-DmfEntity -Lines @([pscustomobject]@{ EntityName = 'Units'; TargetEntity = $null }) -Session (New-TestSession) -Map $m)
        $r[0].Status | Should -Be 'Unresolved-Ambiguous'
        $r[0].Reason | Should -BeLike '*UnitEntityV1*UnitEntityV2*'
    }

    It 'prefers the DataManagementEnabled candidate' {
        $global:MetaLabels['@SYS5'] = 'Currencies'   # a second, non-DM entity shares the label
        try {
            $m = Get-DmfEntityMap -Path (Join-Path $tempRoot 'r4.json')
            $r = @(Resolve-DmfEntity -Lines @([pscustomobject]@{ EntityName = 'Currencies'; TargetEntity = $null }) -Session (New-TestSession) -Map $m)
            $r[0].Status       | Should -Be 'Resolved'
            $r[0].TargetEntity | Should -Be 'CurrencyEntity'
        } finally { $global:MetaLabels['@SYS5'] = 'Internal only' }
    }

    It 'resolves by name candidates and key-segment label reads when the service refuses label filters (the real behaviour)' {
        $global:MetaMock.ValueFilter = $false
        $m = Get-DmfEntityMap -Path (Join-Path $tempRoot 'r5.json')
        $lines = @(
            [pscustomobject]@{ EntityName = 'Customer groups'; TargetEntity = $null }
            [pscustomobject]@{ EntityName = 'Currencies';      TargetEntity = $null }
            [pscustomobject]@{ EntityName = 'Workflow notes';  TargetEntity = $null }
            [pscustomobject]@{ EntityName = 'Nothing like it'; TargetEntity = $null }
        )
        $r = @(Resolve-DmfEntity -Lines $lines -Session (New-TestSession) -Map $m)
        $r[0].Status | Should -Be 'Resolved';  $r[0].Collection | Should -Be 'CustomerGroups'; $r[0].TargetEntity | Should -Be 'CustCustomerGroupEntity'
        $r[1].Status | Should -Be 'Resolved';  $r[1].Collection | Should -Be 'Currencies'
        $r[2].Status | Should -Be 'NotPublic'; $r[2].TargetEntity | Should -Be 'WorkflowNoteEntity'
        $r[3].Status | Should -Be 'Unresolved'; $r[3].Reason | Should -BeLike '*SeedFromPath*'
        # one refused Value-filter attempt, one full DataEntities fetch, no or-chained label batches, only key-segment label reads
        @($global:MetaMock.Calls | Where-Object { $_ -like "*Value eq*" }).Count | Should -Be 1
        @($global:MetaMock.Calls | Where-Object { $_ -like '*/Metadata/DataEntities?$select*' -or $_ -like '*/Metadata/DataEntities' }).Count | Should -Be 1
        @($global:MetaMock.Calls | Where-Object { $_ -like "*Labels?*Id eq*" }).Count | Should -Be 0
        @($global:MetaMock.Calls | Where-Object { $_ -like "*Labels(Id=*" }).Count | Should -BeGreaterThan 0
        @($global:MetaMock.Calls | Where-Object { $_ -like "*Labels(Id=*" }).Count | Should -BeLessThan 20
    }

    It 'flags an ambiguous label through the candidate path too' {
        $global:MetaMock.ValueFilter = $false
        $m = Get-DmfEntityMap -Path (Join-Path $tempRoot 'r7.json')
        $r = @(Resolve-DmfEntity -Lines @([pscustomobject]@{ EntityName = 'Units'; TargetEntity = $null }) -Session (New-TestSession) -Map $m)
        $r[0].Status | Should -Be 'Unresolved-Ambiguous'
    }

    It 'honours a manual entry even with -Refresh' {
        $m = Get-DmfEntityMap -Path (Join-Path $tempRoot 'r6.json')
        [void](Set-DmfEntityMapEntry -Map $m -EntityName 'Units' -Values @{ targetEntity = 'UnitEntityV2'; publicCollectionName = 'UnitsV2'; dataServiceEnabled = $true; keyFields = @('UnitId'); source = 'manual' })
        $r = @(Resolve-DmfEntity -Lines @([pscustomobject]@{ EntityName = 'Units'; TargetEntity = $null }) -Session (New-TestSession) -Map $m -Refresh)
        $r[0].Status     | Should -Be 'Resolved'
        $r[0].Collection | Should -Be 'UnitsV2'
        $r[0].Source     | Should -Be 'manual'
        $global:MetaMock.Calls.Count | Should -Be 0
    }
}
