#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
    Unit tests for lib/DmfAuth.ps1.  No network: Invoke-RestMethod is mocked.
    Run from the repo root:  Invoke-Pester ./tests
#>
Set-StrictMode -Version Latest

BeforeAll {
    $libPath = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'lib'
    . (Join-Path $libPath 'DmfOutput.ps1')
    . (Join-Path $libPath 'DmfRequest.ps1')
    . (Join-Path $libPath 'DmfAuth.ps1')

    $Script:LineWidth  = 80
    $Script:Version    = 'test'
    $Script:MaxRetries = 0      # no transient retries inside tests

    function New-TestSession {
        param([double]$ExpiresInMinutes = 30, [string]$RefreshToken = 'rt-0')
        [pscustomobject]@{
            PSTypeName      = 'Dmf.Session'
            BaseUrl         = 'https://contoso.operations.dynamics.com'
            TenantId        = 'contoso.onmicrosoft.com'
            ClientId        = 'client'
            AuthBase        = 'https://login.microsoftonline.com/contoso.onmicrosoft.com/oauth2/v2.0'
            Scope           = 'https://contoso.operations.dynamics.com/.default offline_access'
            EnvironmentName = 'contoso'
            AccessToken     = 'at-0'
            RefreshToken    = $RefreshToken
            ExpiresAt       = (Get-Date).AddMinutes($ExpiresInMinutes)
            ScopeNote       = $null
            RefreshCount    = 0
            LastRefreshAt   = $null
        }
    }
}

Describe 'Get-DmfEnvironmentName' {
    It 'derives <expected> from <url>' -TestCases @(
        @{ url = 'https://contoso-uat.sandbox.operations.dynamics.com'; expected = 'contoso-uat' }
        @{ url = 'https://contoso.operations.dynamics.com/';            expected = 'contoso' }
        @{ url = 'https://usnconeboxax1aos.cloud.onebox.dynamics.com';  expected = 'usnconeboxax1aos' }
        @{ url = 'https://10.0.0.5';                                    expected = '10-0-0-5' }
        @{ url = 'https://Contoso-PROD.operations.dynamics.com';        expected = 'contoso-prod' }
    ) {
        Get-DmfEnvironmentName -EnvironmentUrl $url | Should -Be $expected
    }
}

Describe 'Get-DmfAuthHeaders' {
    BeforeEach {
        Mock Write-Warn   {}
        Mock Write-Info   {}
        Mock Write-Detail {}
        Mock Invoke-RestMethod { [pscustomobject]@{ access_token = 'at-1'; expires_in = 3600; refresh_token = 'rt-1' } }
    }

    It 'returns the current token without refreshing when plenty of time remains' {
        $s = New-TestSession -ExpiresInMinutes 30
        (Get-DmfAuthHeaders -Session $s).Authorization | Should -Be 'Bearer at-0'
        Should -Invoke Invoke-RestMethod -Times 0 -Exactly
    }

    It 'refreshes inside the expiry window and rotates the refresh token' {
        $s = New-TestSession -ExpiresInMinutes 2
        (Get-DmfAuthHeaders -Session $s).Authorization | Should -Be 'Bearer at-1'
        $s.RefreshCount  | Should -Be 1
        $s.RefreshToken  | Should -Be 'rt-1'
        $s.ExpiresAt     | Should -BeGreaterThan (Get-Date).AddMinutes(50)
        Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter {
            $Uri -like '*/token' -and $Body -like 'grant_type=refresh_token*' -and $Body -notlike '*device_code*'
        }
    }

    It 'does not attempt a refresh when the session has no refresh token' {
        $s = New-TestSession -ExpiresInMinutes 2 -RefreshToken ''
        (Get-DmfAuthHeaders -Session $s).Authorization | Should -Be 'Bearer at-0'
        Should -Invoke Invoke-RestMethod -Times 0 -Exactly
    }

    It 'keeps the current token and stops retrying when the refresh fails' {
        Mock Start-Sleep {}
        Mock Invoke-RestMethod { throw 'AADSTS70008: refresh token expired' }
        $s = New-TestSession -ExpiresInMinutes 2
        (Get-DmfAuthHeaders -Session $s).Authorization | Should -Be 'Bearer at-0'
        $s.RefreshToken | Should -BeNullOrEmpty
        Should -Invoke Write-Warn -Times 1 -Exactly -ParameterFilter { $m -like 'Token refresh failed*' }
        # the refresh allows one transient retry (2 calls); a later header
        # request must not try again because the refresh token was cleared
        Should -Invoke Invoke-RestMethod -Times 2 -Exactly
        [void](Get-DmfAuthHeaders -Session $s)
        Should -Invoke Invoke-RestMethod -Times 2 -Exactly
    }
}

Describe 'Test-DmfTokenExpiry' {
    BeforeEach {
        Mock Write-Warn   {}
        Mock Write-Info   {}
        Mock Write-Detail {}
        Mock Invoke-RestMethod { [pscustomobject]@{ access_token = 'at-1'; expires_in = 3600 } }
    }

    It 'returns false once expired with no refresh token' {
        $s = New-TestSession -ExpiresInMinutes -5 -RefreshToken ''
        Test-DmfTokenExpiry -Session $s | Should -BeFalse
        Should -Invoke Write-Warn -Times 1 -Exactly
    }

    It 'warns but returns true inside the window with no refresh token' {
        $s = New-TestSession -ExpiresInMinutes 2 -RefreshToken ''
        Test-DmfTokenExpiry -Session $s -Activity 'export' | Should -BeTrue
        Should -Invoke Write-Warn -Times 1 -Exactly -ParameterFilter { $m -like '*mid-export*' }
    }

    It 'is silent with plenty of time left' {
        $s = New-TestSession -ExpiresInMinutes 30 -RefreshToken ''
        Test-DmfTokenExpiry -Session $s | Should -BeTrue
        Should -Invoke Write-Warn -Times 0 -Exactly
    }

    It 'refreshes instead of warning when a refresh token is available' {
        $s = New-TestSession -ExpiresInMinutes 2
        Test-DmfTokenExpiry -Session $s | Should -BeTrue
        $s.RefreshCount | Should -Be 1
        Should -Invoke Write-Warn -Times 0 -Exactly
    }
}

Describe 'Connect-DmfEnvironment (device code)' {
    BeforeEach {
        Mock Write-Host   {}
        Mock Write-Info   {}
        Mock Write-Warn   {}
        Mock Write-Detail {}
        Mock Start-Sleep  {}
        Mock Open-DmfBrowser { throw 'a real browser must never open during tests' }
    }

    It 'returns a Dmf.Session with a refresh token and a trimmed base URL' {
        Mock Invoke-RestMethod {
            if ($Uri -like '*/devicecode') { return [pscustomobject]@{ message = 'go'; device_code = 'dc'; expires_in = 60; interval = 0 } }
            if ($Uri -like '*/token')      { return [pscustomobject]@{ access_token = 'at'; expires_in = 3600; refresh_token = 'rt' } }
            throw "unexpected $Uri"
        }
        $s = Connect-DmfEnvironment -EnvironmentUrl 'https://contoso.operations.dynamics.com/' -TenantId 'contoso.onmicrosoft.com' -AuthMode DeviceCode
        $s.AuthMode              | Should -Be 'DeviceCode'
        $s.PSObject.TypeNames[0] | Should -Be 'Dmf.Session'
        $s.BaseUrl               | Should -Be 'https://contoso.operations.dynamics.com'
        $s.EnvironmentName       | Should -Be 'contoso'
        $s.AccessToken           | Should -Be 'at'
        $s.RefreshToken          | Should -Be 'rt'
        $s.Scope                 | Should -Be 'https://contoso.operations.dynamics.com/.default offline_access'
        $s.ScopeNote             | Should -BeNullOrEmpty
        $s.ExpiresAt             | Should -BeGreaterThan (Get-Date).AddMinutes(50)
        Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter { $Uri -like '*/devicecode' -and $Body -like '*offline_access*' }
    }

    It 'requests only the resource scope with -NoOfflineAccess' {
        Mock Invoke-RestMethod {
            if ($Uri -like '*/devicecode') { return [pscustomobject]@{ message = 'go'; device_code = 'dc'; expires_in = 60; interval = 0 } }
            return [pscustomobject]@{ access_token = 'at'; expires_in = 3600 }
        }
        $s = Connect-DmfEnvironment -EnvironmentUrl 'https://contoso.operations.dynamics.com' -TenantId 't' -NoOfflineAccess -AuthMode DeviceCode
        $s.Scope        | Should -Be 'https://contoso.operations.dynamics.com/.default'
        $s.RefreshToken | Should -BeNullOrEmpty
        Should -Invoke Invoke-RestMethod -Times 0 -Exactly -ParameterFilter { $Body -like '*offline_access*' }
    }

    It 'falls back to the resource scope when offline_access is rejected' {
        Mock Invoke-RestMethod {
            if ($Uri -like '*/devicecode' -and $Body -like '*offline_access*') { throw 'AADSTS65001: consent required' }
            if ($Uri -like '*/devicecode') { return [pscustomobject]@{ message = 'go'; device_code = 'dc'; expires_in = 60; interval = 0 } }
            return [pscustomobject]@{ access_token = 'at'; expires_in = 3600 }
        }
        $s = Connect-DmfEnvironment -EnvironmentUrl 'https://contoso.operations.dynamics.com' -TenantId 't' -AuthMode DeviceCode
        $s.Scope        | Should -Be 'https://contoso.operations.dynamics.com/.default'
        $s.RefreshToken | Should -BeNullOrEmpty
        $s.ScopeNote    | Should -Match 'offline_access rejected'
        Should -Invoke Write-Warn -Times 1 -Exactly
        Should -Invoke Invoke-RestMethod -Times 2 -Exactly -ParameterFilter { $Uri -like '*/devicecode' }
    }

    It 'keeps polling through authorization_pending' {
        $global:DmfTestTokenCalls = 0
        Mock Invoke-RestMethod {
            if ($Uri -like '*/devicecode') { return [pscustomobject]@{ message = 'go'; device_code = 'dc'; expires_in = 60; interval = 0 } }
            $global:DmfTestTokenCalls++
            if ($global:DmfTestTokenCalls -lt 3) {
                $er = [System.Management.Automation.ErrorRecord]::new(
                    [System.Exception]::new('HTTP 400'), 'pending', [System.Management.Automation.ErrorCategory]::InvalidOperation, $null)
                $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":"authorization_pending"}')
                throw $er
            }
            return [pscustomobject]@{ access_token = 'at'; expires_in = 3600; refresh_token = 'rt' }
        }
        $s = Connect-DmfEnvironment -EnvironmentUrl 'https://contoso.operations.dynamics.com' -TenantId 't' -AuthMode DeviceCode
        $s.AccessToken            | Should -Be 'at'
        $global:DmfTestTokenCalls | Should -Be 3
    }

    It 'throws when the user declines' {
        Mock Invoke-RestMethod {
            if ($Uri -like '*/devicecode') { return [pscustomobject]@{ message = 'go'; device_code = 'dc'; expires_in = 60; interval = 0 } }
            $er = [System.Management.Automation.ErrorRecord]::new(
                [System.Exception]::new('HTTP 400'), 'declined', [System.Management.Automation.ErrorCategory]::InvalidOperation, $null)
            $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":"authorization_declined"}')
            throw $er
        }
        { Connect-DmfEnvironment -EnvironmentUrl 'https://contoso.operations.dynamics.com' -TenantId 't' -AuthMode DeviceCode } |
            Should -Throw '*Sign-in declined*'
    }
}

Describe 'Connect-DmfEnvironment (browser)' {
    BeforeEach {
        Mock Write-Host   {}
        Mock Write-Info   {}
        Mock Write-Warn   {}
        Mock Write-Detail {}
        Mock Start-Sleep  {}
        Mock New-DmfLoopbackListener { [pscustomobject]@{ Listener = $null; Port = 54321; RedirectUri = 'http://localhost:54321' } }
        Mock Open-DmfBrowser {}
        Mock Wait-DmfAuthRedirect { 'auth-code-1' }
        Mock Invoke-RestMethod {
            if ($Uri -like '*/token' -and $Body -like 'grant_type=authorization_code*') {
                return [pscustomobject]@{ access_token = 'at-b'; expires_in = 3600; refresh_token = 'rt-b'; scope = 'x' }
            }
            if ($Uri -like '*/devicecode') { return [pscustomobject]@{ message = 'go'; device_code = 'dc'; expires_in = 60; interval = 0 } }
            if ($Uri -like '*/token')      { return [pscustomobject]@{ access_token = 'at-d'; expires_in = 3600; refresh_token = 'rt-d' } }
            throw "unexpected $Uri"
        }
    }

    It 'signs in through the browser with PKCE and the loopback redirect' {
        $s = Connect-DmfEnvironment -EnvironmentUrl 'https://contoso.operations.dynamics.com' -TenantId 't' -AuthMode Browser
        $s.AuthMode     | Should -Be 'Browser'
        $s.AccessToken  | Should -Be 'at-b'
        $s.RefreshToken | Should -Be 'rt-b'
        $s.Scope        | Should -Be 'https://contoso.operations.dynamics.com/.default offline_access'
        Should -Invoke Open-DmfBrowser -Times 1 -Exactly -ParameterFilter {
            $Url -like 'https://login.microsoftonline.com/t/oauth2/v2.0/authorize?*' -and
            $Url -like '*client_id=1950a258-227b-4e31-a9cf-717495945fc2*' -and
            $Url -like '*redirect_uri=http%3A%2F%2Flocalhost%3A54321*' -and
            $Url -like '*code_challenge_method=S256*' -and
            $Url -like '*scope=*offline_access*' -and
            $Url -like '*prompt=select_account*'
        }
        Should -Invoke Wait-DmfAuthRedirect -Times 1 -Exactly -ParameterFilter { $ExpectedState -and $ExpectedState.Length -ge 16 }
        Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter {
            $Body -like 'grant_type=authorization_code*' -and $Body -like '*code=auth-code-1*' -and
            $Body -like '*redirect_uri=http%3A%2F%2Flocalhost%3A54321*' -and $Body -like '*code_verifier=*'
        }
        Should -Invoke Invoke-RestMethod -Times 0 -Exactly -ParameterFilter { $Uri -like '*/devicecode' }
    }

    It 'falls back to the device code in Auto mode when the browser flow fails' {
        Mock Wait-DmfAuthRedirect { throw 'No sign-in completed in the browser within 1 seconds.' }
        $s = Connect-DmfEnvironment -EnvironmentUrl 'https://contoso.operations.dynamics.com' -TenantId 't' -AuthMode Auto
        $s.AuthMode    | Should -Be 'DeviceCode'
        $s.AccessToken | Should -Be 'at-d'
        Should -Invoke Write-Warn -Times 1 -Exactly -ParameterFilter { $m -like 'Browser sign-in did not complete*' }
        Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter { $Uri -like '*/devicecode' }
    }

    It 'does not fall back when Browser was requested explicitly' {
        Mock Wait-DmfAuthRedirect { throw 'Sign-in declined in the browser.' }
        { Connect-DmfEnvironment -EnvironmentUrl 'https://contoso.operations.dynamics.com' -TenantId 't' -AuthMode Browser } |
            Should -Throw '*declined*'
        Should -Invoke Invoke-RestMethod -Times 0 -Exactly -ParameterFilter { $Uri -like '*/devicecode' }
    }

    It 'never opens a browser in DeviceCode mode' {
        $s = Connect-DmfEnvironment -EnvironmentUrl 'https://contoso.operations.dynamics.com' -TenantId 't' -AuthMode DeviceCode
        $s.AuthMode | Should -Be 'DeviceCode'
        Should -Invoke Open-DmfBrowser -Times 0 -Exactly
        Should -Invoke New-DmfLoopbackListener -Times 0 -Exactly
    }
}

Describe 'PKCE and token helpers' {
    It 'produces an S256 challenge that matches the verifier' {
        $p = New-DmfPkce
        $p.Verifier.Length  | Should -BeGreaterOrEqual 43
        $p.Verifier         | Should -Match '^[A-Za-z0-9_-]+$'
        $p.State            | Should -Match '^[A-Za-z0-9_-]+$'
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $expected = [Convert]::ToBase64String($sha.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($p.Verifier))).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $p.Challenge | Should -Be $expected
        (New-DmfPkce).Verifier | Should -Not -Be $p.Verifier
    }

    It 'reads a claim from a JWT payload and returns null for garbage' {
        $payload = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes('{"upn":"user@contoso.com","name":"Test User"}')).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $jwt = "eyJhbGciOiJub25lIn0.$payload.sig"
        Get-DmfTokenClaim -Token $jwt -Claim 'upn'  | Should -Be 'user@contoso.com'
        Get-DmfTokenClaim -Token $jwt -Claim 'nope' | Should -BeNullOrEmpty
        Get-DmfTokenClaim -Token 'not-a-jwt' -Claim 'upn' | Should -BeNullOrEmpty
        Get-DmfTokenClaim -Token '' -Claim 'upn' | Should -BeNullOrEmpty
    }
}

Describe 'Invoke-DmfRequest session stamping' {
    BeforeEach {
        Mock Write-Warn   {}
        Mock Write-Info   {}
        Mock Write-Detail {}
    }

    It 'overwrites the Authorization header with the session token for environment requests' {
        $Script:DmfSession = New-TestSession -ExpiresInMinutes 30
        $Script:DmfSession.AccessToken = 'fresh'
        Mock Invoke-RestMethod { [pscustomobject]@{ value = @(); auth = $Headers['Authorization'] } }
        $headers = @{ Authorization = 'Bearer stale' }
        $r = Invoke-DmfRequest -Operation 't' -Params @{ Method = 'Get'; Uri = 'https://contoso.operations.dynamics.com/data/Currencies'; Headers = $headers }
        $r.auth                 | Should -Be 'Bearer fresh'
        $headers['Authorization'] | Should -Be 'Bearer fresh'
        $Script:DmfSession = $null
    }

    It 'leaves requests to other hosts and requests without Authorization untouched' {
        $Script:DmfSession = New-TestSession -ExpiresInMinutes 30
        Mock Invoke-RestMethod { [pscustomobject]@{ keys = @($Headers.Keys) } }
        $blob = @{ 'x-ms-blob-type' = 'BlockBlob' }
        $r = Invoke-DmfRequest -Operation 'blob' -Params @{ Method = 'Put'; Uri = 'https://storage.blob.core.windows.net/c/x?sig=1'; Headers = $blob }
        $r.keys | Should -Not -Contain 'Authorization'
        $other = @{ Authorization = 'Bearer other' }
        [void](Invoke-DmfRequest -Operation 'other' -Params @{ Method = 'Get'; Uri = 'https://graph.microsoft.com/v1.0/me'; Headers = $other })
        $other['Authorization'] | Should -Be 'Bearer other'
        $Script:DmfSession = $null
    }

    It 'works when no session has been set' {
        $Script:DmfSession = $null
        Mock Invoke-RestMethod { [pscustomobject]@{ ok = $true } }
        $h = @{ Authorization = 'Bearer x' }
        (Invoke-DmfRequest -Operation 't' -Params @{ Method = 'Get'; Uri = 'https://contoso.operations.dynamics.com/data/X'; Headers = $h }).ok | Should -BeTrue
        $h['Authorization'] | Should -Be 'Bearer x'
    }

    It 'treats HTTP 501 as permanent (no retries)' {
        $Script:DmfSession = $null
        Mock Start-Sleep {}
        Mock Invoke-RestMethod {
            $resp = [pscustomobject]@{ StatusCode = 501; Headers = $null }
            $ex   = [System.Exception]::new('The remote server returned an error: (501) Not Implemented.')
            throw ($ex | Add-Member -MemberType NoteProperty -Name Response -Value $resp -PassThru)
        }
        { Invoke-DmfRequest -Operation 'labels' -Params @{ Method = 'Get'; Uri = 'https://contoso.operations.dynamics.com/Metadata/Labels?$filter=x'; Headers = @{} } } |
            Should -Throw '*HTTP 501*'
        Should -Invoke Invoke-RestMethod -Times 1 -Exactly
        Should -Invoke Start-Sleep -Times 0 -Exactly
    }

    Context '401 recovery' {
        BeforeAll {
            function New-Http401 {
                # An exception shaped like the ones Invoke-RestMethod throws: a
                # Response property whose StatusCode reads as 401.
                $resp = [pscustomobject]@{ StatusCode = 401; Headers = $null }
                $ex   = [System.Exception]::new('The remote server returned an error: (401) Unauthorized.')
                $ex | Add-Member -MemberType NoteProperty -Name Response -Value $resp -PassThru
            }
        }
        BeforeEach {
            $global:DmfTest401Calls = 0
            $Script:DmfSession = New-TestSession -ExpiresInMinutes 30 -RefreshToken 'rt-0'
        }
        AfterEach { $Script:DmfSession = $null }

        It 'renews the token once and retries with the new bearer' {
            Mock Invoke-RestMethod {
                if ($Uri -like '*/token') { return [pscustomobject]@{ access_token = 'at-1'; expires_in = 3600; refresh_token = 'rt-1' } }
                $global:DmfTest401Calls++
                if ($Headers['Authorization'] -ne 'Bearer at-1') { throw (New-Http401) }
                return [pscustomobject]@{ auth = $Headers['Authorization'] }
            }
            $h = @{ Authorization = 'Bearer at-0' }
            $r = Invoke-DmfRequest -Operation 'poll' -Params @{ Method = 'Get'; Uri = 'https://contoso.operations.dynamics.com/data/Jobs'; Headers = $h }
            $r.auth                          | Should -Be 'Bearer at-1'
            $global:DmfTest401Calls          | Should -Be 2
            $Script:DmfSession.RefreshCount  | Should -Be 1
            Should -Invoke Write-Warn -Times 1 -Exactly -ParameterFilter { $m -like '*HTTP 401 - renewing*' }
        }

        It 'gives up after one renewal when the 401 persists' {
            Mock Invoke-RestMethod {
                if ($Uri -like '*/token') { return [pscustomobject]@{ access_token = 'at-1'; expires_in = 3600 } }
                $global:DmfTest401Calls++
                throw (New-Http401)
            }
            { Invoke-DmfRequest -Operation 'poll' -Params @{ Method = 'Get'; Uri = 'https://contoso.operations.dynamics.com/data/Jobs'; Headers = @{ Authorization = 'Bearer at-0' } } } |
                Should -Throw '*HTTP 401 Unauthorized*'
            $global:DmfTest401Calls | Should -Be 2
            Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter { $Uri -like '*/token' }
        }

        It 'does not attempt recovery without a refresh token or for other hosts' {
            $Script:DmfSession.RefreshToken = $null
            Mock Invoke-RestMethod { $global:DmfTest401Calls++; throw (New-Http401) }
            { Invoke-DmfRequest -Operation 'poll' -Params @{ Method = 'Get'; Uri = 'https://contoso.operations.dynamics.com/data/Jobs'; Headers = @{ Authorization = 'Bearer at-0' } } } |
                Should -Throw '*HTTP 401 Unauthorized*'
            $global:DmfTest401Calls | Should -Be 1

            $Script:DmfSession.RefreshToken = 'rt-0'
            $global:DmfTest401Calls = 0
            { Invoke-DmfRequest -Operation 'other' -Params @{ Method = 'Get'; Uri = 'https://graph.microsoft.com/v1.0/me'; Headers = @{ Authorization = 'Bearer g' } } } |
                Should -Throw '*HTTP 401 Unauthorized*'
            $global:DmfTest401Calls | Should -Be 1
            Should -Invoke Invoke-RestMethod -Times 0 -Exactly -ParameterFilter { $Uri -like '*/token' }
        }

        It 'stamps a renewed token on a retry after a transient fault' {
            Mock Start-Sleep {}
            Mock Invoke-RestMethod {
                if ($Uri -like '*/token') { return [pscustomobject]@{ access_token = 'at-1'; expires_in = 3600 } }
                $global:DmfTest401Calls++
                if ($global:DmfTest401Calls -eq 1) {
                    # simulate the token running out during a retry wait
                    $Script:DmfSession.ExpiresAt = (Get-Date).AddMinutes(1)
                    throw 'simulated network fault'
                }
                return [pscustomobject]@{ auth = $Headers['Authorization'] }
            }
            $r = Invoke-DmfRequest -Operation 'poll' -MaxRetries 1 -Params @{ Method = 'Get'; Uri = 'https://contoso.operations.dynamics.com/data/Jobs'; Headers = @{ Authorization = 'Bearer at-0' } }
            $r.auth | Should -Be 'Bearer at-1'
        }
    }
}
