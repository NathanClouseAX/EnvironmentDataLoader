#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
    Unit tests for the retry engine in lib/DmfRequest.ps1: HTTP 429 handling
    (Retry-After in every form, budgets, caps, jittered fallback) and the
    transient/permanent split.  Invoke-RestMethod and Start-Sleep are mocked,
    so no time actually passes.
#>
Set-StrictMode -Version Latest

BeforeAll {
    $libPath = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'lib'
    . (Join-Path $libPath 'DmfOutput.ps1')
    . (Join-Path $libPath 'DmfRequest.ps1')
    $Script:LineWidth = 80
    $Script:Version   = 'test'
    $Script:MaxRetries = 3

    function New-HttpError {
        # Shaped like the exceptions Invoke-RestMethod throws: a Response with
        # StatusCode and Headers.  A hashtable is enough for Get-DmfHeaderValue,
        # whose last-resort path enumerates key/value pairs.
        param([int]$Status, [hashtable]$Headers = @{})
        $resp = [pscustomobject]@{ StatusCode = $Status; Headers = $Headers }
        $ex   = [System.Exception]::new("The remote server returned an error: ($Status).")
        $ex | Add-Member -MemberType NoteProperty -Name Response -Value $resp -PassThru
    }
    function Invoke-Test {
        param([hashtable]$Extra = @{})
        Invoke-DmfRequest -Operation 'test' -Params @{ Method = 'Get'; Uri = 'https://contoso.operations.dynamics.com/data/X'; Headers = @{} } @Extra
    }
}

Describe 'Get-DmfRetryAfterSeconds' {
    It 'reads delta-seconds, an HTTP-date, and the millisecond hint' {
        Get-DmfRetryAfterSeconds -Response ([pscustomobject]@{ Headers = @{ 'Retry-After' = '17' } })  | Should -Be 17
        $date = [DateTimeOffset]::UtcNow.AddSeconds(90).ToString('r')
        $d = Get-DmfRetryAfterSeconds -Response ([pscustomobject]@{ Headers = @{ 'Retry-After' = $date } })
        $d | Should -BeGreaterOrEqual 85
        $d | Should -BeLessOrEqual 91
        Get-DmfRetryAfterSeconds -Response ([pscustomobject]@{ Headers = @{ 'x-ms-retry-after-ms' = '1500' } }) | Should -Be 2
        # the millisecond hint wins when both are present
        Get-DmfRetryAfterSeconds -Response ([pscustomobject]@{ Headers = @{ 'x-ms-retry-after-ms' = '500'; 'Retry-After' = '60' } }) | Should -Be 1
    }
    It 'returns 0 with no usable hint' {
        Get-DmfRetryAfterSeconds -Response ([pscustomobject]@{ Headers = @{} })                         | Should -Be 0
        Get-DmfRetryAfterSeconds -Response ([pscustomobject]@{ Headers = @{ 'Retry-After' = 'soon' } }) | Should -Be 0
        Get-DmfRetryAfterSeconds -Response $null                                                        | Should -Be 0
    }
}

Describe 'HTTP 429 throttling' {
    BeforeEach {
        Mock Write-Warn   {}
        Mock Write-Detail {}
        Mock Start-Sleep  {}
        $global:DmfTestCalls = 0
        $Script:ThrottleMaxRetries     = 6
        $Script:MaxRetryAfterSeconds   = 300
        $Script:MaxThrottleWaitSeconds = 900
    }

    It 'waits exactly the Retry-After the server asked for, then succeeds' {
        Mock Invoke-RestMethod { $global:DmfTestCalls++; if ($global:DmfTestCalls -eq 1) { throw (New-HttpError 429 @{ 'Retry-After' = '7' }) }; [pscustomobject]@{ ok = $true } }
        (Invoke-Test).ok | Should -BeTrue
        $global:DmfTestCalls | Should -Be 2
        Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter { $Seconds -eq 7 }
        Should -Invoke Write-Warn  -Times 1 -Exactly -ParameterFilter { $m -like '*HTTP 429 throttled - server asked for 7s*' }
    }

    It 'honours the millisecond hint used by Power Platform endpoints' {
        Mock Invoke-RestMethod { $global:DmfTestCalls++; if ($global:DmfTestCalls -eq 1) { throw (New-HttpError 429 @{ 'x-ms-retry-after-ms' = '2500' }) }; [pscustomobject]@{ ok = $true } }
        (Invoke-Test).ok | Should -BeTrue
        Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter { $Seconds -eq 3 }
    }

    It 'backs off with jitter when the server gives no hint' {
        Mock Invoke-RestMethod { $global:DmfTestCalls++; if ($global:DmfTestCalls -le 2) { throw (New-HttpError 429) }; [pscustomobject]@{ ok = $true } }
        (Invoke-Test).ok | Should -BeTrue
        $global:DmfTestCalls | Should -Be 3
        Should -Invoke Start-Sleep -Times 2 -Exactly -ParameterFilter { $Seconds -ge 1 -and $Seconds -le 30 }
        Should -Invoke Write-Warn  -Times 2 -Exactly -ParameterFilter { $m -like '*no Retry-After hint, backing off*' }
    }

    It 'caps a single wait at MaxRetryAfterSeconds' {
        $Script:MaxRetryAfterSeconds = 20
        Mock Invoke-RestMethod { $global:DmfTestCalls++; if ($global:DmfTestCalls -eq 1) { throw (New-HttpError 429 @{ 'Retry-After' = '3600' }) }; [pscustomobject]@{ ok = $true } }
        (Invoke-Test).ok | Should -BeTrue
        Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter { $Seconds -eq 20 }
        Should -Invoke Write-Warn  -Times 1 -Exactly -ParameterFilter { $m -like '*server hint capped to 20s*' }
    }

    It 'gives up after ThrottleMaxRetries with a clear message' {
        $Script:ThrottleMaxRetries = 2
        Mock Invoke-RestMethod { $global:DmfTestCalls++; throw (New-HttpError 429 @{ 'Retry-After' = '1' }) }
        { Invoke-Test } | Should -Throw '*Still throttled (HTTP 429) after 2 retries*'
        $global:DmfTestCalls | Should -Be 3
        Should -Invoke Start-Sleep -Times 2 -Exactly
    }

    It 'stops rather than exceed the total throttle-wait budget' {
        $Script:MaxThrottleWaitSeconds = 10
        Mock Invoke-RestMethod { $global:DmfTestCalls++; throw (New-HttpError 429 @{ 'Retry-After' = '8' }) }
        { Invoke-Test } | Should -Throw '*would exceed the 10s total throttle budget*'
        # first wait of 8 s is allowed (8 <= 10); the second (8 + 8 > 10) is refused
        Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter { $Seconds -eq 8 }
        $global:DmfTestCalls | Should -Be 2
    }

    It 'keeps the throttle budget separate from the transient budget' {
        # MaxRetries 0 means no transient retries at all, yet a 429 is still retried
        Mock Invoke-RestMethod { $global:DmfTestCalls++; if ($global:DmfTestCalls -eq 1) { throw (New-HttpError 429 @{ 'Retry-After' = '1' }) }; [pscustomobject]@{ ok = $true } }
        (Invoke-Test -Extra @{ MaxRetries = 0 }).ok | Should -BeTrue
        $global:DmfTestCalls | Should -Be 2
    }

    It 'honours Retry-After on a 503 too, drawing on the transient budget' {
        Mock Invoke-RestMethod { $global:DmfTestCalls++; if ($global:DmfTestCalls -eq 1) { throw (New-HttpError 503 @{ 'Retry-After' = '4' }) }; [pscustomobject]@{ ok = $true } }
        (Invoke-Test).ok | Should -BeTrue
        Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter { $Seconds -eq 4 }
        Should -Invoke Write-Warn  -Times 1 -Exactly -ParameterFilter { $m -like '*HTTP 503 - retrying in 4s (attempt 1 of 3)*' }
    }

    It 'also throttles downloads through Invoke-DmfDownload' {
        Mock Invoke-WebRequest { $global:DmfTestCalls++; if ($global:DmfTestCalls -eq 1) { throw (New-HttpError 429 @{ 'Retry-After' = '2' }) } }
        Mock Test-Path { $false }
        Invoke-DmfDownload -Uri 'https://blob.core.windows.net/x?sig=1' -OutFile (Join-Path $env:TEMP 'dmf-test-download.bin') -Operation 'download'
        $global:DmfTestCalls | Should -Be 2
        Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter { $Seconds -eq 2 }
    }
}

Describe 'Transient versus permanent failures' {
    BeforeEach {
        Mock Write-Warn   {}
        Mock Write-Detail {}
        Mock Start-Sleep  {}
        $global:DmfTestCalls = 0
    }

    It 'retries 5xx and 408 with exponential back-off, then fails' {
        $Script:MaxRetries = 2
        Mock Invoke-RestMethod { $global:DmfTestCalls++; throw (New-HttpError 502) }
        { Invoke-Test } | Should -Throw '*Failed after 2 retries (HTTP 502)*'
        $global:DmfTestCalls | Should -Be 3
        Should -Invoke Start-Sleep -Times 2 -Exactly
        $Script:MaxRetries = 3
    }

    It 'does not retry 400, 403, 404 or 501' -TestCases @(@{ code = 400 }, @{ code = 403 }, @{ code = 404 }, @{ code = 501 }) {
        Mock Invoke-RestMethod { $global:DmfTestCalls++; throw (New-HttpError $code) }
        { Invoke-Test } | Should -Throw "*HTTP ${code}*"
        $global:DmfTestCalls | Should -Be 1
        Should -Invoke Start-Sleep -Times 0 -Exactly
    }
}
