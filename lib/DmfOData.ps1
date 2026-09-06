<#
.SYNOPSIS
    OData helpers for D365 F&O: URL building, literal escaping, and paging.

.DESCRIPTION
    Dot-source this file (after DmfOutput.ps1 and DmfRequest.ps1) to import:

        ConvertTo-DmfODataLiteral  -- escape a string for use inside $filter quotes
        New-DmfODataUri            -- build /data/<Collection>?... with correct escaping
        Get-DmfODataAll            -- follow @odata.nextLink and emit every item
        Get-DmfODataValue          -- the 'value' array of a response, always an array

.NOTES
    Invoke-DmfRequest (DmfRequest.ps1) is used for every call, so retry,
    throttling, and session-based token refresh all apply.
#>

function ConvertTo-DmfODataLiteral {
    <#
    .SYNOPSIS
        Escapes a value for use inside single quotes in an OData $filter.
    .DESCRIPTION
        OData doubles embedded single quotes:  O'Brien  ->  O''Brien
    .OUTPUTS
        [string]
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Value)
    return $Value.Replace("'", "''")
}


function New-DmfODataUri {
    <#
    .SYNOPSIS
        Builds an OData request URL.

    .PARAMETER BaseUrl
        Environment base URL (no trailing slash needed).

    .PARAMETER Collection
        Entity set name, e.g. 'CustomersV3'.  May include a service prefix
        such as 'Metadata/DataEntities' when -ServicePath is 'Metadata'.

    .PARAMETER ServicePath
        'data' (default) or 'Metadata'.

    .PARAMETER Filter
        Raw $filter expression; it is percent-encoded here.

    .PARAMETER Select
        Comma-separated $select list.

    .PARAMETER Top
        $top value; 0 = omit.

    .PARAMETER Count
        Adds $count=true.

    .PARAMETER CrossCompany
        Adds cross-company=true (must accompany a dataAreaId filter to read a
        company other than the caller's default).

    .OUTPUTS
        [string]
    #>
    param(
        [Parameter(Mandatory)][string]$BaseUrl,
        [Parameter(Mandatory)][string]$Collection,
        [ValidateSet('data', 'Metadata')][string]$ServicePath = 'data',
        [string]$Filter,
        [string]$Select,
        [int]$Top = 0,
        [switch]$Count,
        [switch]$CrossCompany
    )

    # .NET Framework (PS 5.1) leaves the apostrophe unescaped while .NET Core
    # (PS 7) encodes it as %27.  Normalise so both hosts produce the same URL
    # in logs and reports; servers accept either form.
    $encode = { param($s) ([System.Uri]::EscapeDataString($s)).Replace("'", '%27') }

    $parts = [System.Collections.Generic.List[string]]::new()
    if ($CrossCompany) { $parts.Add('cross-company=true') }
    if ($Filter)       { $parts.Add('$filter=' + (& $encode $Filter)) }
    if ($Select)       { $parts.Add('$select=' + (& $encode $Select)) }
    if ($Top -gt 0)    { $parts.Add("`$top=$Top") }
    if ($Count)        { $parts.Add('$count=true') }

    $uri = "$($BaseUrl.TrimEnd('/'))/$ServicePath/$Collection"
    if ($parts.Count -gt 0) { $uri += '?' + ($parts -join '&') }
    return $uri
}


function Get-DmfODataValue {
    <#
    .SYNOPSIS
        Emits the items of an OData response's 'value' collection.
    .DESCRIPTION
        Wrap the call in @() -- a zero- or one-row response then still
        supports .Count and [0] under StrictMode on Windows PowerShell 5.1,
        which has no intrinsic .Count on scalars or $null.
    #>
    param([Parameter(Mandatory)][AllowNull()]$Response)
    if ($null -eq $Response) { return }
    # A collection response carries its items in an array named 'value'.  The
    # property lookup is case-insensitive, and a single entity read by key can
    # legitimately have a scalar property called 'Value' (the Metadata
    # service's Labels entity does), so only an array counts as a collection;
    # anything else is one entity and is returned whole.
    $prop = $Response.PSObject.Properties['value']
    if ($null -ne $prop -and ($prop.Value -is [array] -or ($prop.Value -is [System.Collections.IList] -and $prop.Value -isnot [string]))) {
        return $prop.Value
    }
    return $Response
}


function Get-DmfODataAll {
    <#
    .SYNOPSIS
        Fetches every page of an OData collection and emits each item.

    .DESCRIPTION
        Follows @odata.nextLink until exhausted.  Emits items to the pipeline
        as they arrive; collect with @(Get-DmfODataAll ...) for an array.

    .PARAMETER Uri
        First page URL (including any query options).

    .PARAMETER Operation
        Label for log and error messages; the page number is appended.

    .PARAMETER Headers
        Request headers (normally the Authorization header).

    .PARAMETER MaxItems
        Stop after this many items (0 = unlimited).  The page that crosses the
        limit is still fully requested but only the needed items are emitted.

    .PARAMETER OnPage
        Optional scriptblock invoked after each page with two arguments:
        the page number and the cumulative item count.  Used for progress.

    .OUTPUTS
        One object per collection item.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$Operation,
        [Parameter(Mandatory)][hashtable]$Headers,
        [int]$MaxItems = 0,
        [scriptblock]$OnPage
    )

    $page    = 0
    $emitted = 0
    $nextUri = $Uri

    do {
        $page++
        $resp  = Invoke-DmfRequest -Operation "$Operation (page $page)" -Params @{
            Method  = 'Get'
            Uri     = $nextUri
            Headers = $Headers
        }
        $items = @(Get-DmfODataValue -Response $resp)

        foreach ($item in $items) {
            if ($MaxItems -gt 0 -and $emitted -ge $MaxItems) { break }
            Write-Output $item
            $emitted++
        }

        if ($OnPage) { & $OnPage $page $emitted }

        $linkProp = $resp.PSObject.Properties['@odata.nextLink']
        $nextUri  = if ($null -ne $linkProp -and $linkProp.Value) { [string]$linkProp.Value } else { $null }

        if ($MaxItems -gt 0 -and $emitted -ge $MaxItems) { $nextUri = $null }
    } while ($nextUri)
}
