<#
.SYNOPSIS
    REST client helper for D365 F&O DMF API calls with automatic retry.

.DESCRIPTION
    Dot-source this file to import Invoke-DmfRequest and Invoke-DmfDownload
    into your script.

    The following $Script: variables are read by the caller if set:
        $Script:MaxRetries             -- transient retry count      (default 3)
        $Script:ThrottleMaxRetries     -- HTTP 429 retry count       (default 6)
        $Script:MaxRetryAfterSeconds   -- per-wait ceiling, seconds  (default 300)
        $Script:MaxThrottleWaitSeconds -- total 429 wait budget      (default 900)

    Retry behaviour:
      - HTTP 429 (throttling) is treated as its own class with a separate,
        more generous retry budget, so a throttling storm does not consume
        the allowance reserved for genuine transient faults.  The server's
        Retry-After / x-ms-retry-after-ms hint is honoured exactly when
        present -- the server knows its own recovery window better than any
        client-side guess.
      - HTTP 5xx, 408, and network errors are retried with exponential
        back-off (capped at 30 s) plus jitter.
      - HTTP 401 and other 4xx responses are not retried (non-retryable).
      - OData / D365 error detail is extracted from the response body and
        included in thrown error messages.

.NOTES
    Write-Warn and Write-Detail are called for retry notifications -- dot-source
    DmfOutput.ps1 before using this module to ensure those functions exist.
#>

# =============================================================================
#  Retry policy defaults
# =============================================================================
# Resolved per call so a caller can override any of them by setting the
# corresponding $Script: variable before the first request.
$Script:DmfDefaultMaxRetries         = 3
$Script:DmfDefaultThrottleRetries    = 6
$Script:DmfDefaultMaxRetryAfterSecs  = 300
$Script:DmfDefaultMaxThrottleWaitSec = 900


function Get-DmfHeaderValue {
    <#
    .SYNOPSIS
        Reads a single response header value across PS 5.1 and PS 7 exceptions.

    .DESCRIPTION
        PowerShell 5.1 surfaces a WebHeaderCollection, which supports string
        indexing.  PowerShell 7 surfaces HttpResponseHeaders, which does NOT
        expose an indexer -- indexing it throws, which is why header hints were
        silently lost on PS 7 before this helper existed.  TryGetValues is the
        supported accessor there.

    .OUTPUTS
        [string] the first value for the header, or $null when absent.
    #>
    param(
        [Parameter(Mandatory)][AllowNull()]$Headers,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $Headers) { return $null }

    # -- PS 5.1: WebHeaderCollection supports indexing ------------------------
    if ($Headers -is [System.Net.WebHeaderCollection]) {
        try {
            $v = $Headers[$Name]
            if (-not [string]::IsNullOrWhiteSpace($v)) { return $v }
        } catch {}
        return $null
    }

    # -- PS 7: HttpResponseHeaders requires TryGetValues ---------------------
    try {
        $values = $null
        if ($Headers.TryGetValues($Name, [ref]$values)) {
            $first = @($values) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
            if ($first) { return $first }
        }
    } catch {}

    # -- Last resort: enumerate key/value pairs case-insensitively -----------
    try {
        foreach ($pair in $Headers) {
            if ($pair.Key -ieq $Name) {
                $first = @($pair.Value) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
                if ($first) { return $first }
            }
        }
    } catch {}

    return $null
}


function Get-DmfRetryAfterSeconds {
    <#
    .SYNOPSIS
        Extracts the server's requested retry delay from a response.

    .DESCRIPTION
        Checks, in order of precedence:
          1. x-ms-retry-after-ms  -- millisecond hint used by Dataverse and
                                     other Power Platform endpoints.
          2. Retry-After          -- standard header, in either supported form:
                                     delta-seconds ("120") or an HTTP-date
                                     ("Wed, 21 Oct 2015 07:28:00 GMT").

        Both forms are parsed with RetryConditionHeaderValue, the framework's
        HTTP-spec parser, so an absolute date is converted to a delay relative
        to now rather than being discarded.

    .OUTPUTS
        [int] seconds to wait, or 0 when the server gave no usable hint.
    #>
    param([Parameter(Mandatory)][AllowNull()]$Response)

    if ($null -eq $Response) { return 0 }

    $headers = $null
    try { $headers = $Response.Headers } catch { return 0 }
    if ($null -eq $headers) { return 0 }

    # -- 1. Millisecond hint (Dataverse / Power Platform) --------------------
    foreach ($msHeader in @('x-ms-retry-after-ms', 'Retry-After-Ms')) {
        $raw = Get-DmfHeaderValue -Headers $headers -Name $msHeader
        if ($raw) {
            $ms = 0.0
            if ([double]::TryParse($raw,
                    [System.Globalization.NumberStyles]::Float,
                    [System.Globalization.CultureInfo]::InvariantCulture,
                    [ref]$ms) -and $ms -gt 0) {
                return [int][Math]::Ceiling($ms / 1000.0)
            }
        }
    }

    # -- 2. Standard Retry-After (delta-seconds or HTTP-date) ----------------
    $raw = Get-DmfHeaderValue -Headers $headers -Name 'Retry-After'
    if (-not $raw) { return 0 }

    $parsed = $null
    if ([System.Net.Http.Headers.RetryConditionHeaderValue]::TryParse($raw, [ref]$parsed)) {
        if ($null -ne $parsed.Delta) {
            $secs = $parsed.Delta.TotalSeconds
            if ($secs -gt 0) { return [int][Math]::Ceiling($secs) }
            return 0
        }
        if ($null -ne $parsed.Date) {
            $secs = ($parsed.Date - [DateTimeOffset]::UtcNow).TotalSeconds
            if ($secs -gt 0) { return [int][Math]::Ceiling($secs) }
            return 0
        }
    }

    # -- Lenient fallback for servers that send a bare non-conforming number -
    $secs = 0.0
    if ([double]::TryParse($raw,
            [System.Globalization.NumberStyles]::Float,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [ref]$secs) -and $secs -gt 0) {
        return [int][Math]::Ceiling($secs)
    }

    return 0
}


function Get-DmfBackoffSeconds {
    <#
    .SYNOPSIS
        Computes an exponential back-off delay with equal jitter.

    .DESCRIPTION
        Uses the "equal jitter" strategy: half the computed back-off is fixed
        and half is randomised.  This keeps waits bounded below while stopping
        parallel loops (e.g. several templates exporting at once) from
        synchronising their retries and re-throttling the endpoint together.

    .OUTPUTS
        [int] seconds to wait, minimum 1.
    #>
    param(
        [Parameter(Mandatory)][int]$Attempt,
        [int]$CapSeconds = 30
    )

    $base = [Math]::Min($CapSeconds, [Math]::Pow(2, [Math]::Min($Attempt - 1, 10)) * 2)
    $half = $base / 2.0
    $delay = $half + (Get-Random -Minimum 0.0 -Maximum $half)
    return [int][Math]::Max(1, [Math]::Round($delay))
}


function Invoke-DmfWithRetry {
    <#
    .SYNOPSIS
        Core retry engine shared by Invoke-DmfRequest and Invoke-DmfDownload.

    .PARAMETER Action
        Scriptblock performing the request.  Its output is returned on success.

    .PARAMETER BeforeRetry
        Optional scriptblock run before each retry -- used by the download
        helper to clear a partially written file.

    .NOTES
        Transient faults and throttling are counted separately so that a long
        429 storm cannot exhaust the budget reserved for 5xx/network faults.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        [string]$Operation = 'API request',
        [int]$MaxRetries,
        [int]$ThrottleMaxRetries,
        [int]$MaxRetryAfterSeconds,
        [int]$MaxThrottleWaitSeconds,
        [scriptblock]$BeforeRetry
    )

    # -- Resolve policy ------------------------------------------------------
    if (-not $PSBoundParameters.ContainsKey('MaxRetries')) {
        $MaxRetries = if ($null -ne $Script:MaxRetries) { $Script:MaxRetries } else { $Script:DmfDefaultMaxRetries }
    }
    if (-not $PSBoundParameters.ContainsKey('ThrottleMaxRetries')) {
        $ThrottleMaxRetries = if ($null -ne $Script:ThrottleMaxRetries) { $Script:ThrottleMaxRetries } else { $Script:DmfDefaultThrottleRetries }
    }
    if (-not $PSBoundParameters.ContainsKey('MaxRetryAfterSeconds')) {
        $MaxRetryAfterSeconds = if ($null -ne $Script:MaxRetryAfterSeconds) { $Script:MaxRetryAfterSeconds } else { $Script:DmfDefaultMaxRetryAfterSecs }
    }
    if (-not $PSBoundParameters.ContainsKey('MaxThrottleWaitSeconds')) {
        $MaxThrottleWaitSeconds = if ($null -ne $Script:MaxThrottleWaitSeconds) { $Script:MaxThrottleWaitSeconds } else { $Script:DmfDefaultMaxThrottleWaitSec }
    }

    $transientUsed = 0   # 5xx / 408 / network retries consumed
    $throttleUsed  = 0   # HTTP 429 retries consumed
    $throttleWaited = 0  # cumulative seconds spent waiting out throttling

    while ($true) {
        try {
            return (& $Action)
        }
        catch {
            # -- Classify error -----------------------------------------------
            $httpStatus = 0
            $response   = $null
            if ($null -ne $_.Exception.Response) {
                $response = $_.Exception.Response
                try { $httpStatus = [int]$response.StatusCode } catch {}
            }

            # -- Log raw response body ----------------------------------------
            if ($null -ne $_.ErrorDetails -and $_.ErrorDetails.Message) {
                Write-Detail "[$Operation] Response body: $($_.ErrorDetails.Message)"
            }

            # -- Extract OData / D365 error detail ----------------------------
            $detail = ''
            try {
                $errBody = $_.ErrorDetails.Message | ConvertFrom-Json
                if ($errBody.error.message) {
                    $detail = $errBody.error.message
                    try {
                        $inner = $errBody.error.innererror.message
                        if ($inner) { $detail += " -- $inner" }
                    } catch {}
                }
                elseif ($errBody.'odata.error'.message.value) { $detail = $errBody.'odata.error'.message.value }
            } catch {}
            if (-not $detail) { $detail = $_.Exception.Message }

            # -- Non-retryable client errors ----------------------------------
            if ($httpStatus -eq 401) {
                throw "[$Operation] HTTP 401 Unauthorized. The access token has expired or lacks permission. Re-run to re-authenticate."
            }
            if ($httpStatus -ge 400 -and $httpStatus -lt 500 -and $httpStatus -notin @(408, 429)) {
                throw "[$Operation] HTTP ${httpStatus}: $detail"
            }

            # =================================================================
            #  HTTP 429 -- throttling
            # =================================================================
            if ($httpStatus -eq 429) {
                $serverWait = Get-DmfRetryAfterSeconds -Response $response

                if ($throttleUsed -ge $ThrottleMaxRetries) {
                    throw "[$Operation] Still throttled (HTTP 429) after $ThrottleMaxRetries retries and ${throttleWaited}s of waiting. Reduce concurrency or retry later. Server detail: $detail"
                }

                # Honour the server's hint exactly; only fall back to a guess
                # when it declined to give one.
                if ($serverWait -gt 0) {
                    $delay  = $serverWait
                    $source = "server asked for ${delay}s"
                } else {
                    $delay  = Get-DmfBackoffSeconds -Attempt ($throttleUsed + 1)
                    $source = "no Retry-After hint, backing off ${delay}s"
                }

                # Clamp a single wait so a stray huge hint cannot hang the run.
                if ($delay -gt $MaxRetryAfterSeconds) {
                    Write-Detail "[$Operation] Retry-After of ${delay}s exceeds the ${MaxRetryAfterSeconds}s ceiling; capping."
                    $delay  = $MaxRetryAfterSeconds
                    $source = "server hint capped to ${delay}s"
                }

                # Bail out rather than waiting past the total throttle budget.
                if (($throttleWaited + $delay) -gt $MaxThrottleWaitSeconds) {
                    throw "[$Operation] Throttled (HTTP 429) and the next wait of ${delay}s would exceed the ${MaxThrottleWaitSeconds}s total throttle budget (${throttleWaited}s already waited). Reduce concurrency or retry later. Server detail: $detail"
                }

                $throttleUsed++
                Write-Warn "[$Operation] HTTP 429 throttled - $source (throttle retry $throttleUsed of $ThrottleMaxRetries; ${throttleWaited}s waited so far)..."

                if ($BeforeRetry) { & $BeforeRetry }
                Start-Sleep -Seconds $delay
                $throttleWaited += $delay
                continue
            }

            # =================================================================
            #  5xx / 408 / network -- transient
            # =================================================================
            $label = if ($httpStatus -gt 0) { "HTTP $httpStatus" } else { 'network error' }

            if ($transientUsed -ge $MaxRetries) {
                throw "[$Operation] Failed after $MaxRetries retries ($label): $detail"
            }

            # A 503 may also carry Retry-After; respect it when offered.
            $serverWait = Get-DmfRetryAfterSeconds -Response $response
            $delay = if ($serverWait -gt 0) {
                [Math]::Min($serverWait, $MaxRetryAfterSeconds)
            } else {
                Get-DmfBackoffSeconds -Attempt ($transientUsed + 1)
            }

            $transientUsed++
            Write-Warn "[$Operation] $label - retrying in ${delay}s (attempt $transientUsed of $MaxRetries)..."

            if ($BeforeRetry) { & $BeforeRetry }
            Start-Sleep -Seconds $delay
        }
    }
}


function Invoke-DmfRequest {
    <#
    .SYNOPSIS
        Invokes a REST method with automatic retry and D365-aware error handling.

    .PARAMETER Params
        Hashtable of parameters passed directly to Invoke-RestMethod.

    .PARAMETER Operation
        Human-readable name for this call, used in log and error messages.

    .PARAMETER MaxRetries
        Override the script-level default transient retry count for this call.

    .PARAMETER ThrottleMaxRetries
        Override the script-level default HTTP 429 retry count for this call.

    .OUTPUTS
        The object returned by Invoke-RestMethod on success.

    .EXAMPLE
        $resp = Invoke-DmfRequest -Operation 'GetAzureWriteUrl' -Params @{
            Method      = 'Post'
            Uri         = "$dmfBase.GetAzureWriteUrl"
            Headers     = $authHeaders
            ContentType = 'application/json'
            Body        = (@{ uniqueFileName = $name } | ConvertTo-Json)
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable]$Params,
        [string]$Operation = 'API request',
        [int]$MaxRetries,
        [int]$ThrottleMaxRetries
    )

    Write-Detail "[$Operation] $($Params.Method) $($Params.Uri)"
    if ($Params['ContentType'] -eq 'application/json' -and $Params['Body']) {
        Write-Detail "[$Operation] Body: $($Params['Body'])"
    }

    $forward = @{
        Action    = { Invoke-RestMethod @Params }.GetNewClosure()
        Operation = $Operation
    }
    if ($PSBoundParameters.ContainsKey('MaxRetries'))         { $forward.MaxRetries         = $MaxRetries }
    if ($PSBoundParameters.ContainsKey('ThrottleMaxRetries')) { $forward.ThrottleMaxRetries = $ThrottleMaxRetries }

    Invoke-DmfWithRetry @forward
}


function Invoke-DmfDownload {
    <#
    .SYNOPSIS
        Downloads a file with the same retry and throttling policy as the API calls.

    .DESCRIPTION
        Package downloads go to Azure blob storage, which throttles independently
        of the D365 API and can return 429 or 503 under load.  Routing them
        through the shared retry engine means a throttled download backs off and
        resumes instead of failing the export outright.

        The PowerShell progress bar is suppressed for the duration --
        Invoke-WebRequest renders progress on every packet by default, which
        makes large blob downloads orders of magnitude slower on PS 5.1.

    .PARAMETER Uri
        Source URL to download.

    .PARAMETER OutFile
        Destination file path.

    .PARAMETER Operation
        Human-readable name for this call, used in log and error messages.

    .EXAMPLE
        Invoke-DmfDownload -Uri $downloadUrl -OutFile $zipFile -Operation 'download package'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$OutFile,
        [string]$Operation = 'download file',
        [int]$MaxRetries,
        [int]$ThrottleMaxRetries
    )

    Write-Detail "[$Operation] GET $Uri"

    $action = {
        $prevPref = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing
        } finally {
            $ProgressPreference = $prevPref
        }
    }.GetNewClosure()

    # A failed attempt can leave a truncated file behind; clear it so the
    # retry starts clean and a partial download is never mistaken for success.
    $cleanup = {
        if (Test-Path -LiteralPath $OutFile) {
            try { Remove-Item -LiteralPath $OutFile -Force -ErrorAction Stop } catch {}
        }
    }.GetNewClosure()

    $forward = @{
        Action      = $action
        BeforeRetry = $cleanup
        Operation   = $Operation
    }
    if ($PSBoundParameters.ContainsKey('MaxRetries'))         { $forward.MaxRetries         = $MaxRetries }
    if ($PSBoundParameters.ContainsKey('ThrottleMaxRetries')) { $forward.ThrottleMaxRetries = $ThrottleMaxRetries }

    Invoke-DmfWithRetry @forward
}
