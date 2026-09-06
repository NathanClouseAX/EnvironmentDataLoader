<#
.SYNOPSIS
    Microsoft Entra sign-in and token lifecycle for the D365 F&O scripts.

.DESCRIPTION
    Dot-source this file (after DmfOutput.ps1 and DmfRequest.ps1) to import:

        Connect-DmfEnvironment   -- sign in (browser pop-up or device code); returns a session
        Get-DmfAuthHeaders       -- @{ Authorization = 'Bearer ...' }, refreshing first
                                    when the token is close to expiry
        Test-DmfTokenExpiry      -- warns (or refreshes) as expiry approaches;
                                    returns $false only when the token is dead
        Update-DmfSessionToken   -- refresh_token grant (called by the two above)
        Get-DmfEnvironmentName   -- folder-safe name derived from the URL host
        Get-DmfTokenClaim        -- read a claim (upn, name) from a JWT for display

    Sign-in modes  (-AuthMode)
    ──────────────────────────
      Browser     Authorization-code flow with PKCE.  The default browser opens
                  on the Entra sign-in page (usually already signed in), the
                  user picks an account, and Entra redirects to a loopback
                  listener on http://localhost:<random port> which hands the
                  code back to the script.  No app registration is needed: the
                  public Azure CLI client permits localhost redirects.
      DeviceCode  Prints a code and URL to enter in any browser (RFC 8628).
                  Works over SSH, in containers, and on servers without a
                  browser.
      Auto        Browser when the session is interactive, otherwise device
                  code; a browser attempt that fails or times out falls back
                  to the device code automatically.  (default)

    Silent refresh
    ──────────────
    Both modes request 'offline_access'.  When Entra grants it, the session
    carries a refresh token and Get-DmfAuthHeaders renews the access token
    once fewer than $Script:DmfRefreshWindowMinutes remain.  Invoke-DmfRequest
    (DmfRequest.ps1) calls it on every attempt for requests that carry an
    Authorization header and target the session's environment, and recovers
    from an unexpected HTTP 401 by renewing once.  Scripts opt in with:

        $Script:DmfSession = $session

    Tokens are held in memory only.  Never write AccessToken or RefreshToken
    to logs, transcripts, or files.

.NOTES
    Write-Step / Write-Info / Write-Warn / Format-Elapsed come from
    DmfOutput.ps1 and Invoke-DmfRequest / Invoke-DmfWithRetry /
    Get-DmfRetryAfterSeconds from DmfRequest.ps1 -- dot-source both first.
#>

# Public Azure CLI application -- no app registration required.
$Script:DmfDefaultClientId      = '1950a258-227b-4e31-a9cf-717495945fc2'
# Refresh proactively when fewer than this many minutes remain.  Wider than
# the longest poll interval any script allows (-PollIntervalSeconds max 300),
# so a token cannot slip from "fine" to "expired" between two checks.
$Script:DmfRefreshWindowMinutes = 10
# How long the browser flow waits for the user before giving up.
$Script:DmfBrowserTimeoutSeconds = 180


function Get-DmfEnvironmentName {
    <#
    .SYNOPSIS
        Derives a folder-safe environment name from an environment URL.

    .DESCRIPTION
        Takes the host's first DNS label, lower-cased, and replaces any
        character that is not valid in a file name (plus '.') with '-'.

          https://contoso-uat.sandbox.operations.dynamics.com  -> contoso-uat
          https://contoso.operations.dynamics.com              -> contoso
          https://usnconeboxax1aos.cloud.onebox.dynamics.com   -> usnconeboxax1aos
          https://10.0.0.5                                     -> 10-0-0-5

    .OUTPUTS
        [string]
    #>
    param([Parameter(Mandatory)][string]$EnvironmentUrl)

    $uri = [System.Uri]$EnvironmentUrl
    $label = $uri.Host
    if ($uri.HostNameType -ne [System.UriHostNameType]::IPv4 -and
        $uri.HostNameType -ne [System.UriHostNameType]::IPv6) {
        $label = ($label -split '\.')[0]
    }
    $label = $label.ToLowerInvariant()

    $invalid = [System.IO.Path]::GetInvalidFileNameChars() + @([char]'.', [char]':', [char]'[', [char]']')
    $sb = [System.Text.StringBuilder]::new()
    foreach ($ch in $label.ToCharArray()) {
        if ($invalid -contains $ch) { [void]$sb.Append('-') } else { [void]$sb.Append($ch) }
    }
    $name = $sb.ToString().Trim('-')
    if ([string]::IsNullOrWhiteSpace($name)) { $name = 'environment' }
    return $name
}


function Get-DmfTokenClaim {
    <#
    .SYNOPSIS
        Reads one claim from a JWT payload for display purposes (no signature check).
    .OUTPUTS
        [string] or $null
    #>
    param([Parameter(Mandatory)][AllowNull()][AllowEmptyString()][string]$Token, [Parameter(Mandatory)][string]$Claim)
    try {
        if ([string]::IsNullOrEmpty($Token)) { return $null }
        $parts = $Token.Split('.')
        if ($parts.Count -lt 2) { return $null }
        $payload = $parts[1].Replace('-', '+').Replace('_', '/')
        switch ($payload.Length % 4) { 2 { $payload += '==' } 3 { $payload += '=' } }
        $json = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
        $p = $json.PSObject.Properties[$Claim]
        if ($null -ne $p -and $null -ne $p.Value) { return [string]$p.Value }
    } catch {}
    return $null
}


function ConvertTo-DmfBase64Url {
    param([Parameter(Mandatory)][byte[]]$Bytes)
    return [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}


function New-DmfPkce {
    <#
    .SYNOPSIS  Generates a PKCE verifier / S256 challenge pair and a state nonce.
    #>
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $vBytes = New-Object byte[] 32
        $rng.GetBytes($vBytes)
        $sBytes = New-Object byte[] 16
        $rng.GetBytes($sBytes)
    } finally { $rng.Dispose() }
    $verifier  = ConvertTo-DmfBase64Url -Bytes $vBytes
    $sha       = [System.Security.Cryptography.SHA256]::Create()
    try { $challenge = ConvertTo-DmfBase64Url -Bytes $sha.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($verifier)) }
    finally { $sha.Dispose() }
    return [pscustomobject]@{ Verifier = $verifier; Challenge = $challenge; State = (ConvertTo-DmfBase64Url -Bytes $sBytes) }
}


function New-DmfLoopbackListener {
    <#
    .SYNOPSIS
        Starts a TCP listener on 127.0.0.1 at a free port for the OAuth redirect.
    .DESCRIPTION
        A raw TcpListener rather than HttpListener: it needs no URL ACL
        reservation and therefore no administrator rights.
    .OUTPUTS
        [pscustomobject]  Listener, Port, RedirectUri
    #>
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
    $listener.Start()
    $port = ([System.Net.IPEndPoint]$listener.LocalEndpoint).Port
    return [pscustomobject]@{ Listener = $listener; Port = $port; RedirectUri = "http://localhost:$port" }
}


function Open-DmfBrowser {
    <#
    .SYNOPSIS  Opens a URL in the default browser.
    #>
    param([Parameter(Mandatory)][string]$Url)
    Start-Process -FilePath $Url | Out-Null
}


function Wait-DmfAuthRedirect {
    <#
    .SYNOPSIS
        Waits for Entra to redirect the browser to the loopback listener and
        returns the authorization code.

    .DESCRIPTION
        Accepts connections until one carries a query string with 'code' or
        'error' (browsers also request /favicon.ico), answers it with a small
        HTML page, validates the state nonce, and returns the code.  Throws on
        timeout, on an error response from Entra, or on a state mismatch.
    #>
    param(
        [Parameter(Mandatory)]$Loopback,
        [Parameter(Mandatory)][string]$ExpectedState,
        [int]$TimeoutSeconds = $Script:DmfBrowserTimeoutSeconds
    )

    $listener = $Loopback.Listener
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $page = {
        param($title, $body)
        "<!DOCTYPE html><html><head><meta charset='utf-8'><title>$title</title></head>" +
        "<body style='font-family:Segoe UI,system-ui,sans-serif;background:#f3f2f1;color:#323130;display:flex;align-items:center;justify-content:center;height:100vh;margin:0'>" +
        "<div style='background:#fff;padding:32px 40px;border-radius:4px;box-shadow:0 1px 3px rgba(0,0,0,.1);max-width:520px'>" +
        "<h1 style='font-size:20px;margin:0 0 8px'>$title</h1><p style='margin:0'>$body</p></div></body></html>"
    }

    try {
        while ((Get-Date) -lt $deadline) {
            $ar = $listener.BeginAcceptTcpClient($null, $null)
            $remaining = [Math]::Max(1, [int](($deadline - (Get-Date)).TotalMilliseconds))
            if (-not $ar.AsyncWaitHandle.WaitOne($remaining)) { break }
            $client = $listener.EndAcceptTcpClient($ar)
            try {
                $client.ReceiveTimeout = 5000
                $stream = $client.GetStream()
                $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::ASCII)
                $requestLine = $reader.ReadLine()
                # Drain the headers so the browser sees a clean close.
                while ($true) { $l = $reader.ReadLine(); if ($null -eq $l -or $l -eq '') { break } }

                $query = @{}
                if ($requestLine -match '^GET\s+[^?\s]*\?([^\s]*)\s+HTTP') {
                    foreach ($pair in ($Matches[1] -split '&')) {
                        $kv = $pair -split '=', 2
                        $k  = [System.Uri]::UnescapeDataString($kv[0])
                        $v  = if ($kv.Count -gt 1) { [System.Uri]::UnescapeDataString($kv[1].Replace('+', ' ')) } else { '' }
                        $query[$k] = $v
                    }
                }

                $isAuthResponse = $query.ContainsKey('code') -or $query.ContainsKey('error')
                $html = if (-not $isAuthResponse) { & $page 'Waiting for sign-in' 'This window is waiting for Microsoft Entra to complete sign-in.' }
                        elseif ($query.ContainsKey('error')) { & $page 'Sign-in failed' "$($query['error']): $($query['error_description'])  You can close this window." }
                        else { & $page 'Signed in' 'You can close this window and return to the PowerShell script.' }
                $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($html)
                $status    = if ($isAuthResponse) { '200 OK' } else { '404 Not Found' }
                $header    = "HTTP/1.1 $status`r`nContent-Type: text/html; charset=utf-8`r`nContent-Length: $($bodyBytes.Length)`r`nConnection: close`r`n`r`n"
                $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($header)
                $stream.Write($headerBytes, 0, $headerBytes.Length)
                $stream.Write($bodyBytes, 0, $bodyBytes.Length)
                $stream.Flush()

                if (-not $isAuthResponse) { continue }
                if ($query.ContainsKey('error')) {
                    if ($query['error'] -eq 'access_denied') { throw 'Sign-in declined in the browser.' }
                    throw "Entra returned '$($query['error'])': $($query['error_description'])"
                }
                if ($query['state'] -ne $ExpectedState) { throw 'Sign-in response state did not match the request (possible interference); aborting.' }
                return [string]$query['code']
            }
            finally { $client.Close() }
        }
        throw "No sign-in completed in the browser within $TimeoutSeconds seconds."
    }
    finally {
        try { $listener.Stop() } catch {}
    }
}


function Invoke-DmfBrowserSignIn {
    <#
    .SYNOPSIS
        Authorization-code + PKCE sign-in through the default browser.
    .OUTPUTS
        The token endpoint response (access_token, expires_in, refresh_token, ...).
    #>
    param(
        [Parameter(Mandatory)][string]$AuthBase,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$Scope,
        [int]$TimeoutSeconds = $Script:DmfBrowserTimeoutSeconds
    )

    $pkce     = New-DmfPkce
    $loopback = New-DmfLoopbackListener
    $authUrl  = "$AuthBase/authorize?" + (@(
        "client_id=$ClientId",
        'response_type=code',
        "redirect_uri=$([System.Uri]::EscapeDataString($loopback.RedirectUri))",
        "scope=$([System.Uri]::EscapeDataString($Scope))",
        "state=$($pkce.State)",
        "code_challenge=$($pkce.Challenge)",
        'code_challenge_method=S256',
        'prompt=select_account'
    ) -join '&')

    Write-Info "Opening your browser to sign in (listening on $($loopback.RedirectUri); waiting up to ${TimeoutSeconds}s)."
    Write-Detail 'If no browser window appears, re-run with -AuthMode DeviceCode.'
    Open-DmfBrowser -Url $authUrl
    $code = Wait-DmfAuthRedirect -Loopback $loopback -ExpectedState $pkce.State -TimeoutSeconds $TimeoutSeconds

    # Exchange the code.  Sent through Invoke-DmfWithRetry rather than
    # Invoke-DmfRequest so the body (which carries the code) is never logged.
    $body = "grant_type=authorization_code&client_id=$ClientId" +
            "&code=$([System.Uri]::EscapeDataString($code))" +
            "&redirect_uri=$([System.Uri]::EscapeDataString($loopback.RedirectUri))" +
            "&code_verifier=$($pkce.Verifier)" +
            "&scope=$([System.Uri]::EscapeDataString($Scope))"
    return Invoke-DmfWithRetry -Operation 'authorization code exchange' -MaxRetries 1 -ArgumentList @("$AuthBase/token", $body) -Action {
        param($tokenUri, $formBody)
        Invoke-RestMethod -Method Post -Uri $tokenUri -ContentType 'application/x-www-form-urlencoded' -Body $formBody
    }
}


function Invoke-DmfDeviceCodeSignIn {
    <#
    .SYNOPSIS
        RFC 8628 device-code sign-in: prints the code and URL, polls the token endpoint.
    .OUTPUTS
        [pscustomobject]  Token (endpoint response), Scope (as granted), ScopeNote
    #>
    param(
        [Parameter(Mandatory)][string]$AuthBase,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$Scope,
        [Parameter(Mandatory)][string]$ResourceScope
    )

    $scopeNote = $null
    $deviceCode = $null
    try {
        $deviceCode = Invoke-DmfRequest -Operation 'device code request' -Params @{
            Method      = 'Post'
            Uri         = "$AuthBase/devicecode"
            ContentType = 'application/x-www-form-urlencoded'
            Body        = "client_id=$ClientId&scope=$([System.Uri]::EscapeDataString($Scope))"
        }
    }
    catch {
        if ($Scope -eq $ResourceScope) { throw }
        # Tenant policy may reject the extra scope; fall back to the plain
        # resource scope so the script still runs, just without silent refresh.
        $scopeNote = "offline_access rejected at /devicecode: $($_.Exception.Message)"
        Write-Warn "Could not request offline_access ($($_.Exception.Message)); continuing without silent token refresh."
        $Scope = $ResourceScope
        $deviceCode = Invoke-DmfRequest -Operation 'device code request (resource scope only)' -Params @{
            Method      = 'Post'
            Uri         = "$AuthBase/devicecode"
            ContentType = 'application/x-www-form-urlencoded'
            Body        = "client_id=$ClientId&scope=$([System.Uri]::EscapeDataString($Scope))"
        }
    }

    Write-Host ''
    Write-Host $deviceCode.message -ForegroundColor Yellow
    Write-Host ''

    $pollUntil    = (Get-Date).AddSeconds([int]$deviceCode.expires_in)
    $pollInterval = [int]$deviceCode.interval
    $tokenResp    = $null

    while ((Get-Date) -lt $pollUntil) {
        if ($pollInterval -gt 0) { Start-Sleep -Seconds $pollInterval }
        try {
            $tokenResp = Invoke-RestMethod -Method Post `
                -Uri         "$AuthBase/token" `
                -ContentType 'application/x-www-form-urlencoded' `
                -Body        "grant_type=urn:ietf:params:oauth:grant-type:device_code&client_id=$ClientId&device_code=$($deviceCode.device_code)"
            break
        }
        catch {
            $errCode = $null
            try { $errCode = ($_.ErrorDetails.Message | ConvertFrom-Json).error } catch {}

            $status   = 0
            $response = $null
            # Only HTTP exceptions carry a Response property; a network fault does
            # not, and a direct read would throw under StrictMode and mask it.
            $responseProp = $_.Exception.PSObject.Properties['Response']
            if ($null -ne $responseProp -and $null -ne $responseProp.Value) {
                $response = $responseProp.Value
                try { $status = [int]$response.StatusCode } catch {}
            }

            # Throttled polling is routine, not a failure.  RFC 8628 requires the
            # client to lengthen its interval by 5 s on slow_down; a 429 on the
            # token endpoint is treated the same way, honouring Retry-After when
            # one is supplied.  Handled outside the switch below because
            # 'continue' inside a switch continues the switch, not the loop.
            if ($status -eq 429 -or $errCode -eq 'slow_down') {
                $wait         = Get-DmfRetryAfterSeconds -Response $response
                $pollInterval = if ($wait -gt 0) { $wait } else { $pollInterval + 5 }
                Write-Warn "Sign-in polling throttled - slowing to ${pollInterval}s between checks..."
            }
            else {
                switch ($errCode) {
                    'authorization_pending'  { continue }
                    'authorization_declined' { throw 'Sign-in declined.  Re-run and approve the prompt.' }
                    'expired_token'          { throw 'Device code expired.  Re-run the script.' }
                    default                  { throw }
                }
            }
        }
    }

    if ($null -eq $tokenResp) { throw 'Authentication timed out before sign-in completed.' }
    return [pscustomobject]@{ Token = $tokenResp; Scope = $Scope; ScopeNote = $scopeNote }
}


function Connect-DmfEnvironment {
    <#
    .SYNOPSIS
        Signs in to a D365 F&O environment and returns a session object.

    .PARAMETER EnvironmentUrl
        Base URL of the environment.  A trailing slash is removed.

    .PARAMETER TenantId
        Entra tenant ID or domain.

    .PARAMETER ClientId
        Application (client) ID.  Defaults to the public Azure CLI client.

    .PARAMETER AuthMode
        Browser, DeviceCode, or Auto (default; see the file header).

    .PARAMETER NoOfflineAccess
        Request only the resource scope; no refresh token.

    .PARAMETER BrowserTimeoutSeconds
        How long the browser flow waits for the user (default 180).

    .OUTPUTS
        [pscustomobject] Dmf.Session
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidatePattern('^https?://')][string]$EnvironmentUrl,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$TenantId,
        [string]$ClientId = $Script:DmfDefaultClientId,
        [ValidateSet('Auto', 'Browser', 'DeviceCode')][string]$AuthMode = 'Auto',
        [switch]$NoOfflineAccess,
        [int]$BrowserTimeoutSeconds = $Script:DmfBrowserTimeoutSeconds
    )

    $baseUrl       = $EnvironmentUrl.TrimEnd('/')
    $authBase      = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0"
    $resourceScope = "$baseUrl/.default"
    $scope         = if ($NoOfflineAccess) { $resourceScope } else { "$resourceScope offline_access" }
    $scopeNote     = $null
    $tokenResp     = $null
    $modeUsed      = 'DeviceCode'

    $tryBrowser = ($AuthMode -eq 'Browser') -or ($AuthMode -eq 'Auto' -and [Environment]::UserInteractive)
    if ($tryBrowser) {
        try {
            $tokenResp = Invoke-DmfBrowserSignIn -AuthBase $authBase -ClientId $ClientId -Scope $scope -TimeoutSeconds $BrowserTimeoutSeconds
            $modeUsed  = 'Browser'
        }
        catch {
            if ($AuthMode -eq 'Browser') { throw }
            Write-Warn "Browser sign-in did not complete ($($_.Exception.Message)); falling back to the device code flow."
            $tokenResp = $null
        }
    }

    if ($null -eq $tokenResp) {
        $dc        = Invoke-DmfDeviceCodeSignIn -AuthBase $authBase -ClientId $ClientId -Scope $scope -ResourceScope $resourceScope
        $tokenResp = $dc.Token
        $scope     = $dc.Scope
        $scopeNote = $dc.ScopeNote
    }

    $accessToken  = [string]$tokenResp.access_token
    if ([string]::IsNullOrEmpty($accessToken)) { throw 'Sign-in completed but no access token was returned.' }
    $expiresAt    = (Get-Date).AddSeconds([int]$tokenResp.expires_in - 60)   # 60 s safety buffer
    $refreshToken = $null
    $rtProp = $tokenResp.PSObject.Properties['refresh_token']
    if ($null -ne $rtProp -and -not [string]::IsNullOrEmpty($rtProp.Value)) { $refreshToken = [string]$rtProp.Value }

    $session = [pscustomobject]@{
        PSTypeName      = 'Dmf.Session'
        BaseUrl         = $baseUrl
        TenantId        = $TenantId
        ClientId        = $ClientId
        AuthBase        = $authBase
        AuthMode        = $modeUsed
        Scope           = $scope
        EnvironmentName = Get-DmfEnvironmentName -EnvironmentUrl $baseUrl
        AccessToken     = $accessToken
        RefreshToken    = $refreshToken
        ExpiresAt       = $expiresAt
        ScopeNote       = $scopeNote
        RefreshCount    = 0
        LastRefreshAt   = $null
    }

    $who = Get-DmfTokenClaim -Token $accessToken -Claim 'upn'
    if (-not $who) { $who = Get-DmfTokenClaim -Token $accessToken -Claim 'unique_name' }
    if (-not $who) { $who = Get-DmfTokenClaim -Token $accessToken -Claim 'preferred_username' }
    $lifetime   = Format-Elapsed ($expiresAt - (Get-Date))
    $refreshTxt = if ($refreshToken) { 'Silent refresh enabled -- runs longer than this will renew the token automatically.' }
                  else { 'No refresh token was issued -- a run longer than this will need re-authentication.' }
    Write-Info "Sign-in successful$(if ($who) { " as $who" }) via $modeUsed.  Access token valid for $lifetime (until $($expiresAt.ToString('HH:mm:ss'))).  $refreshTxt"

    return $session
}


function Update-DmfSessionToken {
    <#
    .SYNOPSIS
        Renews the session's access token with the refresh_token grant.

    .DESCRIPTION
        Updates AccessToken, ExpiresAt and (when Entra rotates it)
        RefreshToken on the session in place.  The token endpoint is called
        through Invoke-DmfWithRetry rather than Invoke-DmfRequest so the
        request body -- which contains the refresh token -- is never logged.

    .OUTPUTS
        [bool]  $true when the token was renewed.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][PSTypeName('Dmf.Session')]$Session)

    if ([string]::IsNullOrEmpty($Session.RefreshToken)) { return $false }

    $authBase = $Session.AuthBase
    $body     = "grant_type=refresh_token&client_id=$($Session.ClientId)" +
                "&refresh_token=$([System.Uri]::EscapeDataString($Session.RefreshToken))" +
                "&scope=$([System.Uri]::EscapeDataString($Session.Scope))"

    Write-Detail '[token refresh] Renewing access token with refresh_token grant.'
    $resp = Invoke-DmfWithRetry -Operation 'token refresh' -MaxRetries 1 -ArgumentList @("$authBase/token", $body) -Action {
        param($tokenUri, $formBody)
        Invoke-RestMethod -Method Post -Uri $tokenUri `
            -ContentType 'application/x-www-form-urlencoded' -Body $formBody
    }

    $Session.AccessToken = $resp.access_token
    $Session.ExpiresAt   = (Get-Date).AddSeconds([int]$resp.expires_in - 60)
    $rtProp = $resp.PSObject.Properties['refresh_token']
    if ($null -ne $rtProp -and -not [string]::IsNullOrEmpty($rtProp.Value)) { $Session.RefreshToken = [string]$rtProp.Value }
    $Session.RefreshCount++
    $Session.LastRefreshAt = Get-Date

    Write-Info "Access token renewed; now valid until $($Session.ExpiresAt.ToString('HH:mm:ss'))."
    return $true
}


function Get-DmfAuthHeaders {
    <#
    .SYNOPSIS
        Returns the Authorization header for the session, refreshing first if
        the token is about to expire.

    .OUTPUTS
        [hashtable]  @{ Authorization = 'Bearer <token>' }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][PSTypeName('Dmf.Session')]$Session)

    $window = (Get-Date).AddMinutes($Script:DmfRefreshWindowMinutes)
    if ($window -ge $Session.ExpiresAt -and -not [string]::IsNullOrEmpty($Session.RefreshToken)) {
        try {
            [void](Update-DmfSessionToken -Session $Session)
        } catch {
            Write-Warn "Token refresh failed ($($_.Exception.Message)); continuing with the current token."
            # Do not retry on every call once refresh has failed.
            $Session.RefreshToken = $null
        }
    }
    return @{ Authorization = "Bearer $($Session.AccessToken)" }
}


function Test-DmfTokenExpiry {
    <#
    .SYNOPSIS
        Warns as the token nears expiry, refreshing it instead when possible.

    .PARAMETER Activity
        Word used in the warning text, e.g. 'export' or 'import'.

    .OUTPUTS
        [bool]  $false when the token has expired and cannot be renewed;
                $true otherwise.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSTypeName('Dmf.Session')]$Session,
        [string]$Activity = 'run'
    )

    $now = Get-Date
    if (-not [string]::IsNullOrEmpty($Session.RefreshToken)) {
        # Get-DmfAuthHeaders performs the refresh when inside the window.
        [void](Get-DmfAuthHeaders -Session $Session)
        return ($now -lt $Session.ExpiresAt)
    }

    if ($now -ge $Session.ExpiresAt) {
        Write-Warn 'Access token has expired.  API calls will likely fail with HTTP 401.  Re-run the script.'
        return $false
    }
    if ($now.AddMinutes($Script:DmfRefreshWindowMinutes) -ge $Session.ExpiresAt) {
        Write-Warn "Token expires at $($Session.ExpiresAt.ToString('HH:mm:ss')) -- it may expire mid-$Activity."
    }
    return $true
}
