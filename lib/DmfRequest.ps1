<#
.SYNOPSIS
    REST client helper for D365 F&O DMF API calls with automatic retry.

.DESCRIPTION
    Dot-source this file to import Invoke-DmfRequest into your script.

    The following $Script: variable must be set by the caller:
        $Script:MaxRetries -- default retry count (range 0-10)

    Retry behaviour:
      - HTTP 5xx, 408, 429, and network errors are retried with exponential
        back-off (capped at 30 s).  The Retry-After response header is
        honoured when present.
      - HTTP 401 and other 4xx responses are not retried (non-retryable).
      - OData / D365 error detail is extracted from the response body and
        included in thrown error messages.

.NOTES
    Write-Warn is called for retry notifications -- dot-source DmfOutput.ps1
    before using this module to ensure that function is available.
#>

function Invoke-DmfRequest {
    <#
    .SYNOPSIS
        Invokes a REST method with automatic retry and D365-aware error handling.

    .PARAMETER Params
        Hashtable of parameters passed directly to Invoke-RestMethod.

    .PARAMETER Operation
        Human-readable name for this call, used in log and error messages.

    .PARAMETER MaxRetries
        Override the script-level default retry count for this call.

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
        [int]$MaxRetries   = $Script:MaxRetries
    )

    Write-Verbose "[$Operation] $($Params.Method) $($Params.Uri)"

    for ($attempt = 1; $attempt -le ($MaxRetries + 1); $attempt++) {
        try {
            return Invoke-RestMethod @Params
        }
        catch {
            # -- Classify error -----------------------------------------------
            $httpStatus = 0
            $retryAfter = 0
            if ($null -ne $_.Exception.Response) {
                try { $httpStatus = [int]$_.Exception.Response.StatusCode } catch {}
                try {
                    $ra = $_.Exception.Response.Headers['Retry-After']
                    if ($ra) { $retryAfter = [int]$ra }
                } catch {}
            }

            # -- Extract OData / D365 error detail ----------------------------
            $detail = ''
            try {
                $errBody = $_.ErrorDetails.Message | ConvertFrom-Json
                if ($errBody.error.message)                   { $detail = $errBody.error.message }
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

            # -- Exhausted retries --------------------------------------------
            if ($attempt -gt $MaxRetries) {
                $label = if ($httpStatus -gt 0) { "HTTP $httpStatus" } else { 'network error' }
                throw "[$Operation] Failed after $MaxRetries retries ($label): $detail"
            }

            # -- Retryable: back off then retry -------------------------------
            $delay = if ($retryAfter -gt 0) { $retryAfter } else {
                [Math]::Min(30, [Math]::Pow(2, $attempt - 1) * 2)
            }
            $label = if ($httpStatus -gt 0) { "HTTP $httpStatus" } else { 'network error' }
            Write-Warn "[$Operation] $label - retrying in ${delay}s (attempt $attempt of $MaxRetries)..."
            Start-Sleep -Seconds $delay
        }
    }
}
