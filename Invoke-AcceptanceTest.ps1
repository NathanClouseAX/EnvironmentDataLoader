#Requires -Version 5.1
<#
.SYNOPSIS
    End-to-end acceptance test: runs every script in this repo in every mode,
    in dependency order, checks each outcome, and writes a report.

.DESCRIPTION
    Each step runs the real script in a child PowerShell process (the scripts
    call 'exit', so they cannot be dot-sourced), captures its console output
    to a per-step log, collects its -PassThru objects, and asserts on the
    result.  Later steps consume earlier ones: the template captured in
    'capture' drives the exports and the OData pull; the DMF project created
    by 'dmf-export-local' is what 'template-export' exports; that zip is what
    'expand' extracts and 'import-whatif' validates; the OData folders are
    what 'compare' diffs.  A failed step does not stop the run; steps whose
    inputs are missing are skipped and say why.

    Offline steps (no sign-in, no environment):
      parse            every .ps1 parses on this host
      pester           Invoke-Pester ./tests
      whatif           -WhatIf of every script that supports it
      compare-fixtures Compare-EnvironmentData.ps1 on tests/fixtures (all options)

    Online steps (sign in once per step -- with -AuthMode Auto the browser
    opens and you pick the account; expect one prompt per step):
      probe            Invoke-EnvironmentProbe.ps1
      capture          Export-TemplateDefinition.ps1 -TemplateName X into <WorkPath>\resources
      seed             Export-TemplateDefinition.ps1 -SeedFromPath (after an export exists)
      capture-all      Export-TemplateDefinition.ps1 -All          (only with -Full)
      dmf-export-local Invoke-ProjectExport.ps1 -TemplateSource Local -Mode Dmf
      dmf-export-env   Invoke-ProjectExport.ps1 -TemplateSource Environment -Mode Dmf
      template-export  Invoke-TemplateExport.ps1 against the project the export created
      expand           Expand-ExportedPackages.ps1 on that zip
      import-whatif    Invoke-BaselineImport.ps1 -WhatIf on the extracted package
      upload-whatif    Invoke-PackageUpload.ps1 -WhatIf on that zip
      odata-pull       Invoke-ProjectExport.ps1 -Mode OData (local template)
      odata-pull-env   Invoke-ProjectExport.ps1 -Mode OData -TemplateSource Environment
      odata-pull-2     the same pull for -SecondLegalEntityId (if given)
      compare          Compare-EnvironmentData.ps1 on the pulled folders (all options)
      auth-devicecode  one sign-in with -AuthMode DeviceCode          (only with -TestDeviceCode)
      import           Invoke-BaselineImport.ps1 for real              (only with -AllowImport, into -ImportEnvironmentUrl)
      upload           Invoke-PackageUpload.ps1 for real               (only with -AllowImport)
      job-report       Get-ExecutionJobReport.ps1 on the import environment (after 'import')

    Nothing is written to the source environment except the DMF export
    project that Invoke-ProjectExport.ps1 always creates ("<template> <LE>").
    Imports and uploads are WhatIf unless -AllowImport is given together with
    a sandbox -ImportEnvironmentUrl.

.PARAMETER EnvironmentUrl
    Source environment for the online steps.  Omit with -Offline.

.PARAMETER TenantId
    Entra tenant.  Omit with -Offline.

.PARAMETER LegalEntityId
    Company for exports and pulls.

.PARAMETER SecondLegalEntityId
    Optional second company: pulled too and diffed against the first with
    -IgnoreFields dataAreaId.

.PARAMETER TemplateName
    Template used throughout.  Default '010 - System Setup'.

.PARAMETER WorkPath
    Scratch root for resources/, downloads/, data/, logs/ and the report.
    Default: $env:TEMP\DmfAcceptance_<timestamp>.  The repo's own resources/
    and data/ are never touched.

.PARAMETER Steps
    Step names (wildcards) to run.  Default: all applicable.

.PARAMETER SkipSteps
    Step names (wildcards) to skip.

.PARAMETER Offline
    Run only the steps that need no environment.

.PARAMETER Full
    Include the long steps (capture-all).

.PARAMETER TestDeviceCode
    Add one sign-in with -AuthMode DeviceCode (you will have to type a code).

.PARAMETER AllowImport
    Run the real import and upload into -ImportEnvironmentUrl / -ImportLegalEntityId.
    Never point this at production.

.PARAMETER ImportEnvironmentUrl
    Sandbox environment for -AllowImport.  Default: -EnvironmentUrl (refused
    unless -IReallyMeanTheSourceEnvironment is also given).

.PARAMETER ImportLegalEntityId
    Company for -AllowImport.  Default: -LegalEntityId.

.PARAMETER AuthMode
    Passed to every online step.  Default Auto.

.PARAMETER PollIntervalSeconds
    Passed to the DMF export/import steps.  Default 10.

.PARAMETER ScriptHost
    'pwsh' or 'powershell' for the child processes.  Default: the host
    running this script.

.PARAMETER WhatIf
    Print the planned steps and their command lines; run nothing.

.PARAMETER PassThru
    Emit one result object per step.

.EXAMPLE
    # Everything that needs no environment
    .\Invoke-AcceptanceTest.ps1 -Offline

.EXAMPLE
    # Full online run against a sandbox, imports included
    .\Invoke-AcceptanceTest.ps1 -EnvironmentUrl 'https://contoso-uat.sandbox.operations.dynamics.com' -TenantId 'contoso.onmicrosoft.com' `
        -LegalEntityId 'USMF' -SecondLegalEntityId 'DAT' -AllowImport -ImportEnvironmentUrl 'https://contoso-test.sandbox.operations.dynamics.com'

.EXAMPLE
    # Only the OData path and the diff
    .\Invoke-AcceptanceTest.ps1 -EnvironmentUrl 'https://x.operations.dynamics.com' -TenantId 'x' -LegalEntityId 'DAT' -Steps 'capture', 'odata-*', 'compare'
#>
[CmdletBinding()]
param(
    [ValidatePattern('^https?://')]
    [string]$EnvironmentUrl,

    [string]$TenantId,

    [string]$LegalEntityId,

    [string]$SecondLegalEntityId,

    [string]$TemplateName = '010 - System Setup',

    [string]$WorkPath,

    [string[]]$Steps,
    [string[]]$SkipSteps,

    [switch]$Offline,
    [switch]$Full,
    [switch]$TestDeviceCode,

    [switch]$AllowImport,
    [ValidatePattern('^(https?://.*)?$')]
    [string]$ImportEnvironmentUrl,
    [string]$ImportLegalEntityId,
    [switch]$IReallyMeanTheSourceEnvironment,

    [ValidateSet('Auto', 'Browser', 'DeviceCode')]
    [string]$AuthMode = 'Auto',

    [ValidateRange(5, 300)]
    [int]$PollIntervalSeconds = 10,

    [ValidateSet('pwsh', 'powershell')]
    [string]$ScriptHost,

    [switch]$WhatIf,
    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = $PSScriptRoot
. (Join-Path $repo 'lib\DmfOutput.ps1')
$Script:Version   = '1.0'
$Script:LineWidth = 80
try { $Script:LineWidth = [Math]::Min(120, [Math]::Max(80, $Host.UI.RawUI.BufferSize.Width)) } catch {}

# =============================================================================
#  Pre-flight
# =============================================================================
if (-not $Offline) {
    foreach ($p in 'EnvironmentUrl', 'TenantId', 'LegalEntityId') {
        if (-not (Get-Variable -Name $p -ValueOnly)) { throw "-$p is required unless -Offline is given." }
    }
}
if ($AllowImport) {
    if (-not $ImportEnvironmentUrl) { $ImportEnvironmentUrl = $EnvironmentUrl }
    if (-not $ImportLegalEntityId)  { $ImportLegalEntityId  = $LegalEntityId }
    if ($ImportEnvironmentUrl -eq $EnvironmentUrl -and -not $IReallyMeanTheSourceEnvironment) {
        throw "-AllowImport would import into the source environment '$EnvironmentUrl'.  Give a sandbox -ImportEnvironmentUrl, or add -IReallyMeanTheSourceEnvironment."
    }
}
if (-not $ScriptHost) { $ScriptHost = if ($PSVersionTable.PSVersion.Major -ge 6) { 'pwsh' } else { 'powershell' } }
$hostExe = if ($ScriptHost -eq 'pwsh') { 'pwsh' } else { 'powershell.exe' }

$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
if (-not $WorkPath) { $WorkPath = Join-Path $env:TEMP "DmfAcceptance_$ts" }
$dirs = @{}
foreach ($d in 'resources', 'downloads', 'data', 'logs', 'reports', 'packages') {
    $dirs[$d] = Join-Path $WorkPath $d
    if (-not $WhatIf) { New-Item -ItemType Directory -Path $dirs[$d] -Force | Out-Null }
}
$envName = if ($EnvironmentUrl) { ([System.Uri]$EnvironmentUrl).Host.Split('.')[0].ToLowerInvariant() } else { '' }

# =============================================================================
#  Helpers
# =============================================================================
function ConvertTo-ArgString {
    <# Renders a hashtable of parameters as a PowerShell argument string. #>
    param([hashtable]$Params)
    $parts = foreach ($k in $Params.Keys) {
        $v = $Params[$k]
        if ($v -is [switch] -or $v -is [bool]) { if ($v) { "-$k" }; continue }
        if ($null -eq $v) { continue }
        if ($v -is [array]) { "-$k " + (($v | ForEach-Object { "'" + ([string]$_).Replace("'", "''") + "'" }) -join ','); continue }
        "-$k '" + ([string]$v).Replace("'", "''") + "'"
    }
    return ($parts -join ' ')
}

$results = [System.Collections.Generic.List[pscustomobject]]::new()
$state   = @{}   # values handed from one step to the next
if ($WhatIf) {
    # Placeholders so the plan shows every command line, including the steps
    # that consume an earlier step's output.
    $state['templateFolder']   = '<capture: template folder>'
    $state['dmfPackageFolder'] = '<dmf-export-local: extracted package folder>'
    $state['dmfProjectName']   = "$TemplateName $LegalEntityId"
    $state['zip']              = '<template-export: downloaded zip>'
    $state['expandedFolder']   = '<expand: extracted folder>'
    $state['dataFolder1']      = "<odata-pull: $WorkPath\data\$envName\$LegalEntityId>"
    if ($SecondLegalEntityId) { $state['dataFolder2'] = "<odata-pull-2: $WorkPath\data\$envName\$SecondLegalEntityId>" }
    $state['imported']         = $true
}

function Invoke-Step {
    <#
      Runs one step.  -Script + -Params run a repo script in a child process
      (stdout/stderr to a log, -PassThru objects via CliXml); -Body runs
      in-process instead.  -Check receives (exitCode, passThru, logPath) and
      returns a detail string (throw to fail).  -Requires names $state keys
      that must exist, else the step is skipped.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Title,
        [string]$Script,
        $Params = @{},              # hashtable, or a scriptblock returning one (evaluated after the Requires check)
        [scriptblock]$Body,
        [scriptblock]$Check,
        [string[]]$Requires = @(),
        [switch]$Online,
        [switch]$NoPassThru,
        [int]$TimeoutMinutes = 90
    )

    $r = [pscustomobject]@{ Step = $Name; Title = $Title; Status = 'SKIP'; Detail = ''; Elapsed = '-'; Log = ''; Command = '' }

    $wanted = (-not $Steps) -or (@($Steps | Where-Object { $Name -like $_ }).Count -gt 0)
    $skipped = $SkipSteps -and (@($SkipSteps | Where-Object { $Name -like $_ }).Count -gt 0)
    if (-not $wanted -or $skipped) { $r.Detail = 'not selected'; $results.Add($r); return }
    if ($Online -and $Offline)     { $r.Detail = 'needs an environment (-Offline)'; $results.Add($r); return }
    $missing = @($Requires | Where-Object { -not $state.ContainsKey($_) })
    if ($missing.Count -gt 0)      { $r.Detail = "input from an earlier step missing: $($missing -join ', ')"; $results.Add($r); Write-Warn "$Name skipped -- $($r.Detail)"; return }
    if ($Params -is [scriptblock]) { $Params = & $Params }
    if ($null -eq $Params) { $Params = @{} }

    $cmd = ''
    if ($Script) {
        $outXml = Join-Path $dirs.logs "$Name.passthru.xml"
        $inner  = "& '$(Join-Path $repo $Script)' $(ConvertTo-ArgString $Params)"
        if (-not $NoPassThru) { $inner += " -PassThru | Export-Clixml -Path '$outXml' -Depth 6" }
        $cmd = "$hostExe -NoProfile -ExecutionPolicy Bypass -Command `"$($inner.Replace('"', '\"'))`""
    }
    $r.Command = if ($cmd) { $cmd } else { '<in-process>' }
    $r.Log     = Join-Path $dirs.logs "$Name.log"

    Write-Step "$Name  --  $Title"
    if ($cmd) { Write-Detail $cmd }
    if ($WhatIf) { $r.Status = 'PLANNED'; $results.Add($r); return }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $exit = 0; $pt = $null
        if ($Script) {
            $inner = "& '$(Join-Path $repo $Script)' $(ConvertTo-ArgString $Params)"
            if (-not $NoPassThru) { $inner += " -PassThru | Export-Clixml -Path '$outXml' -Depth 6" }
            $inner += "; exit `$LASTEXITCODE"
            & $hostExe -NoProfile -ExecutionPolicy Bypass -Command $inner *> $r.Log
            $exit = $LASTEXITCODE
            if (Test-Path -LiteralPath $outXml) {
                # Keep only real result objects; a stray empty string or $null on
                # the pipeline must not shift $pt[0].
                try { $pt = @(Import-Clixml -Path $outXml | Where-Object { $null -ne $_ -and -not ($_ -is [string]) }) } catch { $pt = $null }
            }
        } else {
            # In-process bodies return their result on the success stream; child
            # processes they start write their own logs.  Do not redirect here or
            # the result would land in the log instead of $pt.
            $pt = @(& $Body)
            if (-not (Test-Path -LiteralPath $r.Log)) { $r.Log = '' }
        }
        $detail = if ($Check) { [string](& $Check $exit $pt $r.Log) } else { "exit $exit" }
        if ($null -eq $detail) { $detail = "exit $exit" }
        $r.Status = 'PASS'; $r.Detail = $detail
        Write-OK "$Name  $detail"
    }
    catch {
        $r.Status = 'FAIL'; $r.Detail = $_.Exception.Message
        Write-Fail "$Name  $($_.Exception.Message)"
        if (Test-Path -LiteralPath $r.Log) {
            $tail = @(Get-Content -LiteralPath $r.Log -Tail 6 | Where-Object { $_ -match '\S' })
            foreach ($l in $tail) { Write-Detail $l.Substring(0, [Math]::Min(160, $l.Length)) }
        }
    }
    $sw.Stop()
    $r.Elapsed = Format-Elapsed $sw.Elapsed
    $results.Add($r)
}

function Get-LogText { param([string]$Path) if (Test-Path -LiteralPath $Path) { Get-Content -LiteralPath $Path -Raw } else { '' } }
function Assert-LogHas { param([string]$Path, [string]$Pattern, [string]$Because) if ((Get-LogText $Path) -notmatch $Pattern) { throw "expected the log to contain '$Pattern' ($Because)" } }
function Assert-LogLacks { param([string]$Path, [string]$Pattern, [string]$Because) if ((Get-LogText $Path) -match $Pattern) { throw "log contains '$Pattern' ($Because)" } }

# =============================================================================
#  Banner
# =============================================================================
Write-Banner -DryRun:$WhatIf -Title 'EnvironmentDataLoader Acceptance Test'
Write-Host ''
Write-Info "Repo        : $repo"
Write-Info "Work path   : $WorkPath"
Write-Info "Child host  : $hostExe"
if ($Offline) { Write-Info 'Mode        : offline only' }
else {
    Write-Info "Environment : $EnvironmentUrl  ($envName)"
    Write-Info "Tenant      : $TenantId"
    Write-Info "Company     : $LegalEntityId$(if ($SecondLegalEntityId) { " (and $SecondLegalEntityId)" })"
    Write-Info "Template    : $TemplateName"
    Write-Info "Auth mode   : $AuthMode  (expect one sign-in prompt per online step)"
    Write-Info "Imports     : $(if ($AllowImport) { "REAL into $ImportEnvironmentUrl / $ImportLegalEntityId" } else { 'WhatIf only (-AllowImport to run for real)' })"
}
$common = @{ EnvironmentUrl = $EnvironmentUrl; TenantId = $TenantId; AuthMode = $AuthMode; LogPath = '' }

# =============================================================================
#  Offline steps
# =============================================================================
Invoke-Step -Name 'parse' -Title 'every script parses' -NoPassThru -Body {
    $bad = @()
    foreach ($f in (Get-ChildItem -Path $repo -Filter *.ps1 -Recurse | Where-Object { $_.FullName -notmatch '\\\.claude\\' })) {
        $t = $null; $e = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$t, [ref]$e)
        if ($e.Count) { $bad += "$($f.Name): $($e[0].Message)" }
    }
    if ($bad.Count) { throw ($bad -join '; ') }
    "all .ps1 files parse"
} -Check { param($exit, $pt, $log) "$(@(Get-ChildItem -Path $repo -Filter *.ps1 -Recurse | Where-Object { $_.FullName -notmatch '\\\.claude\\' }).Count) files" }

Invoke-Step -Name 'pester' -Title 'unit tests (Invoke-Pester ./tests)' -NoPassThru -Body {
    Import-Module Pester -MinimumVersion 5.0 -ErrorAction Stop
    $res = Invoke-Pester -Path (Join-Path $repo 'tests') -Output None -PassThru
    if ($res.FailedCount -gt 0) { throw "$($res.FailedCount) test(s) failed: " + (($res.Failed | ForEach-Object { $_.ExpandedPath }) -join '; ') }
    $Script:PesterPassed = $res.PassedCount
    "$($res.PassedCount) passed"
} -Check { param($exit, $pt, $log) "$($Script:PesterPassed) tests passed" }

$whatIfEnv = @{ EnvironmentUrl = $(if ($EnvironmentUrl) { $EnvironmentUrl } else { 'https://offline.operations.dynamics.com' }); TenantId = $(if ($TenantId) { $TenantId } else { 'offline.onmicrosoft.com' }); LogPath = '' }
$whatIfLe  = if ($LegalEntityId) { $LegalEntityId } else { 'DAT' }

Invoke-Step -Name 'whatif' -Title '-WhatIf of every script (no sign-in)' -NoPassThru -Body {
    $runs = @(
        @{ S = 'Invoke-ProjectExport.ps1';      P = $whatIfEnv + @{ LegalEntityId = $whatIfLe; TemplateName = '010 - System Setup'; WhatIf = $true } }
        @{ S = 'Invoke-ProjectExport.ps1';      P = $whatIfEnv + @{ LegalEntityId = $whatIfLe; TemplateName = '010 - System Setup'; Mode = 'OData'; WhatIf = $true } }
        @{ S = 'Invoke-ProjectExport.ps1';      P = $whatIfEnv + @{ LegalEntityId = $whatIfLe; TemplateName = '010 - System Setup'; TemplateSource = 'Environment'; WhatIf = $true } }
        @{ S = 'Invoke-TemplateExport.ps1';     P = $whatIfEnv + @{ LegalEntityId = $whatIfLe; TemplateName = '010 - System Setup'; WhatIf = $true } }
        @{ S = 'Invoke-BaselineImport.ps1';     P = $whatIfEnv + @{ LegalEntityId = $whatIfLe; PackageName = '010 - System Setup'; WhatIf = $true } }
        @{ S = 'Export-TemplateDefinition.ps1'; P = $whatIfEnv + @{ TemplateName = 'Acceptance probe'; ResourcesPath = $dirs.resources; WhatIf = $true } }
        @{ S = 'Invoke-EnvironmentProbe.ps1';   P = $whatIfEnv + @{ LegalEntityId = $whatIfLe; OutputPath = ''; WhatIf = $true } }
    )
    $failed = @()
    foreach ($run in $runs) {
        $log = Join-Path $dirs.logs ("whatif." + $run.S.Replace('.ps1', '') + '.' + ($run.P.Keys | Where-Object { $_ -in 'Mode', 'TemplateSource' } | ForEach-Object { $run.P[$_] }) -join '' + '.log')
        $inner = "& '$(Join-Path $repo $run.S)' $(ConvertTo-ArgString $run.P); exit `$LASTEXITCODE"
        & $hostExe -NoProfile -ExecutionPolicy Bypass -Command $inner *> $log
        $text = Get-LogText $log
        if ($LASTEXITCODE -ne 0 -or $text -notmatch 'No API calls made|nothing will be|no changes will be|would be created|No projects created') { $failed += "$($run.S) ($(ConvertTo-ArgString ($run.P | Select-Object -Property *)))" }
        if ($text -match 'Authenticating with Microsoft Entra') { $failed += "$($run.S) tried to sign in under -WhatIf" }
    }
    if ($failed.Count) { throw ($failed -join '; ') }
    "$($runs.Count) WhatIf runs, none signed in"
} -Check { param($exit, $pt, $log) '7 WhatIf runs, none signed in' }

Invoke-Step -Name 'compare-fixtures' -Title 'Compare-EnvironmentData.ps1 on the test fixtures (all options)' -NoPassThru -Body {
    $a = Join-Path $repo 'tests\fixtures\data\env-a\USMF'
    $b = Join-Path $repo 'tests\fixtures\data\env-b\USMF'
    $html = Join-Path $dirs.reports 'compare-fixtures.html'; $json = Join-Path $dirs.reports 'compare-fixtures.json'
    $xml  = Join-Path $dirs.logs 'compare-fixtures.passthru.xml'
    $inner = "& '$(Join-Path $repo 'Compare-EnvironmentData.ps1')' -Folder1 '$a' -Folder2 '$b' -HtmlPath '$html' -JsonPath '$json' -LogPath '' -PassThru | Export-Clixml '$xml' -Depth 6; exit `$LASTEXITCODE"
    # child logs get their own names: the step's own log ('compare-fixtures.log') is
    # held open by the outer redirection while this runs.
    & $hostExe -NoProfile -ExecutionPolicy Bypass -Command $inner *> (Join-Path $dirs.logs 'compare-fixtures.run.log')
    if ($LASTEXITCODE -ne 0) { throw "exit $LASTEXITCODE" }
    $f = @(Import-Clixml $xml)
    $types = $f | Group-Object ChangeType | ForEach-Object { "$($_.Name)=$($_.Count)" }
    if (@($f | Where-Object ChangeType -eq 'Changed').Count -ne 1) { throw "expected 1 Changed finding, got: $($types -join ' ')" }
    if (-not (Test-Path $html) -or -not (Test-Path $json)) { throw 'HTML or JSON not written' }
    $j = Get-Content $json -Raw | ConvertFrom-Json
    if (@($j.templates).Count -ne 3) { throw "expected 3 templates in the JSON roll-up, got $(@($j.templates).Count)" }
    # -Template filter, -ChangesOnly, -Strict, -FailOnDifference exit code, self-compare
    & $hostExe -NoProfile -ExecutionPolicy Bypass -Command "& '$(Join-Path $repo 'Compare-EnvironmentData.ps1')' -Folder1 '$a' -Folder2 '$b' -Template '010*' -ChangesOnly -Strict -HtmlPath '' -LogPath '' -FailOnDifference; exit `$LASTEXITCODE" *> (Join-Path $dirs.logs 'compare-fixtures.template.log')
    if ($LASTEXITCODE -ne 2) { throw "-FailOnDifference should exit 2 when differences exist, got $LASTEXITCODE" }
    & $hostExe -NoProfile -ExecutionPolicy Bypass -Command "& '$(Join-Path $repo 'Compare-EnvironmentData.ps1')' -Folder1 '$a' -Folder2 '$a' -HtmlPath '' -LogPath '' -FailOnDifference; exit `$LASTEXITCODE" *> (Join-Path $dirs.logs 'compare-fixtures.self.log')
    if ($LASTEXITCODE -ne 0) { throw "self-compare should report no differences, exit $LASTEXITCODE" }
    "findings: $($types -join ' '); template filter, ChangesOnly, Strict, FailOnDifference, self-compare OK"
} -Check { param($exit, $pt, $log) $pt[-1] }

# =============================================================================
#  Online steps
# =============================================================================
# The cross-company probe needs a company that is not the caller's default and
# has customer groups; the second company, when given, is the better bet.
$probeLe = if ($SecondLegalEntityId) { $SecondLegalEntityId } else { $LegalEntityId }
Invoke-Step -Name 'probe' -Title "Invoke-EnvironmentProbe.ps1 (cross-company probe on $probeLe)" -Online -Script 'Invoke-EnvironmentProbe.ps1' `
    -Params ($common + @{ LegalEntityId = $probeLe; OutputPath = $dirs.reports }) `
    -Check {
        param($exit, $pt, $log)
        $by = @{}; foreach ($p in $pt) { $by[$p.Id] = $p }
        foreach ($must in 'V1', 'V3', 'V6') { if ($by[$must].Result -ne 'PASS') { throw "$must $($by[$must].Result): $($by[$must].Detail)" } }
        (($pt | ForEach-Object { "$($_.Id)=$($_.Result)" }) -join ' ')
    }

Invoke-Step -Name 'capture' -Title "Export-TemplateDefinition.ps1 -TemplateName '$TemplateName'" -Online -Script 'Export-TemplateDefinition.ps1' `
    -Params ($common + @{ TemplateName = $TemplateName; ResourcesPath = $dirs.resources; Force = $true }) `
    -Check {
        param($exit, $pt, $log)
        $r = $pt | Where-Object TemplateId -eq $TemplateName
        if (-not $r -or $r.Status -ne 'Captured') { throw "status $(if ($r) { $r.Status } else { 'none' })" }
        foreach ($f in 'Manifest.xml', 'PackageHeader.xml', 'template.json') { if (-not (Test-Path (Join-Path $r.Folder $f))) { throw "$f missing" } }
        if (-not (Test-Path (Join-Path $dirs.resources 'entity-map.json'))) { throw 'entity-map.json not written' }
        Assert-LogLacks $log 'HTTP 501 - retrying' 'a 501 must not be retried'
        $state['templateFolder'] = $r.Folder
        if ($r.Resolved -eq 0 -and $r.Unresolved -gt 0) { throw "resolution is not working: 0 of $($r.Unresolved) entities resolved (see $log)" }
        "$($r.Lines) lines, $($r.Resolved) resolved, $($r.Unresolved) unresolved -> $($r.Folder)"
    }

if ($Full) {
    Invoke-Step -Name 'capture-all' -Title 'Export-TemplateDefinition.ps1 -All' -Online -Script 'Export-TemplateDefinition.ps1' -TimeoutMinutes 240 `
        -Params ($common + @{ All = $true; ResourcesPath = $dirs.resources; Force = $true }) `
        -Check { param($exit, $pt, $log) "$(@($pt | Where-Object Status -eq 'Captured').Count) captured, $(@($pt | Where-Object Status -like 'Skipped*').Count) skipped, $(@($pt | Where-Object Status -eq 'Failed').Count) failed of $(@($pt).Count)" }
}

Invoke-Step -Name 'dmf-export-local' -Title 'Invoke-ProjectExport.ps1 -TemplateSource Local -Mode Dmf' -Online -Script 'Invoke-ProjectExport.ps1' -Requires 'templateFolder' `
    -Params ($common + @{ LegalEntityId = $LegalEntityId; TemplateName = $TemplateName; ResourcesPath = $dirs.resources; DownloadPath = $dirs.downloads; PollIntervalSeconds = $PollIntervalSeconds; Force = $true }) `
    -Check {
        param($exit, $pt, $log)
        $r = $pt[0]
        if ($r.Status -notin 'Succeeded', 'PartiallySucceeded') { throw "status $($r.Status) (execution $($r.ExecutionId))" }
        if (-not (Test-Path (Join-Path $r.DownloadedTo 'Manifest.xml'))) { throw "no Manifest.xml in $($r.DownloadedTo)" }
        $xlsx = @(Get-ChildItem $r.DownloadedTo -Filter *.xlsx).Count
        $state['dmfPackageFolder'] = $r.DownloadedTo
        $state['dmfProjectName']   = $r.ProjectName
        "$($r.Status), $($r.LinesAdded) lines, $xlsx xlsx extracted to $($r.DownloadedTo)"
    }

Invoke-Step -Name 'dmf-export-env' -Title 'Invoke-ProjectExport.ps1 -TemplateSource Environment -Mode Dmf' -Online -Script 'Invoke-ProjectExport.ps1' `
    -Params ($common + @{ LegalEntityId = $LegalEntityId; TemplateName = $TemplateName; TemplateSource = 'Environment'; ResourcesPath = $dirs.resources; DownloadPath = $dirs.downloads; PollIntervalSeconds = $PollIntervalSeconds; Force = $true }) `
    -Check {
        param($exit, $pt, $log)
        $r = $pt[0]
        if ($r.Status -notin 'Succeeded', 'PartiallySucceeded') { throw "status $($r.Status)" }
        if (-not $state.ContainsKey('dmfPackageFolder')) { $state['dmfPackageFolder'] = $r.DownloadedTo; $state['dmfProjectName'] = $r.ProjectName }
        "$($r.Status), $($r.LinesAdded) lines -> $($r.DownloadedTo)"
    }

Invoke-Step -Name 'seed' -Title 'Export-TemplateDefinition.ps1 -SeedFromPath (from the exported package)' -Online -Script 'Export-TemplateDefinition.ps1' -Requires 'dmfPackageFolder' -NoPassThru `
    -Params { $common + @{ SeedFromPath = $state['dmfPackageFolder']; ResourcesPath = $dirs.resources } } `
    -Check { param($exit, $pt, $log) Assert-LogHas $log 'Seed complete' 'seed-only run'; Assert-LogLacks $log 'Authenticating' 'seeding needs no sign-in'; ((Get-LogText $log) -split "`n" | Where-Object { $_ -match 'manifest\(s\)' } | Select-Object -First 1).Trim() }

Invoke-Step -Name 'template-export' -Title 'Invoke-TemplateExport.ps1 against the project the export created' -Online -Script 'Invoke-TemplateExport.ps1' -Requires 'dmfProjectName' `
    -Params { $common + @{ LegalEntityId = $LegalEntityId; TemplateName = $state['dmfProjectName']; DownloadPath = $dirs.downloads; PollIntervalSeconds = $PollIntervalSeconds; Force = $true } } `
    -Check {
        param($exit, $pt, $log)
        $r = $pt[0]
        if ($r.Status -notin 'Succeeded', 'PartiallySucceeded') { throw "status $($r.Status)" }
        if (-not (Test-Path $r.DownloadedTo)) { throw "zip not found: $($r.DownloadedTo)" }
        $state['zip'] = $r.DownloadedTo
        "$($r.Status) -> $($r.DownloadedTo)"
    }

Invoke-Step -Name 'expand' -Title 'Expand-ExportedPackages.ps1 on that zip' -Script 'Expand-ExportedPackages.ps1' -Requires 'zip' `
    -Params { @{ SourcePath = $dirs.downloads; DestinationPath = $dirs.packages; PackageName = [System.IO.Path]::GetFileNameWithoutExtension([string]$state['zip']); Force = $true } } `
    -Check {
        param($exit, $pt, $log)
        $r = $pt[0]
        if ($r.Status -ne 'Extracted' -and $r.Status -notlike 'Succe*') { throw "status $($r.Status)" }
        if (-not (Test-Path (Join-Path $r.ExtractedTo 'Manifest.xml'))) { throw 'no Manifest.xml after extraction' }
        $state['expandedFolder'] = $r.ExtractedTo
        "$($r.XlsxCount) xlsx, $($r.EntityCount) entities -> $($r.ExtractedTo)"
    }

Invoke-Step -Name 'import-whatif' -Title 'Invoke-BaselineImport.ps1 -WhatIf on the extracted package' -Online -Script 'Invoke-BaselineImport.ps1' -Requires 'expandedFolder' `
    -Params { $common + @{ LegalEntityId = $LegalEntityId; ResourcesPath = $dirs.packages; PackageName = (Split-Path -Leaf ([string]$state['expandedFolder'])); WhatIf = $true } } `
    -Check { param($exit, $pt, $log) Assert-LogHas $log 'No API calls made' 'WhatIf'; Assert-LogLacks $log 'Authenticating' 'WhatIf must not sign in'; ((Get-LogText $log) -split "`n" | Where-Object { $_ -match 'entities\)' } | Select-Object -First 1).Trim() }

Invoke-Step -Name 'upload-whatif' -Title 'Invoke-PackageUpload.ps1 -WhatIf on that zip' -Online -Script 'Invoke-PackageUpload.ps1' -Requires 'zip' `
    -Params { $common + @{ LegalEntityId = $LegalEntityId; UploadPath = $dirs.downloads; PackageName = (Split-Path -Leaf ([string]$state['zip'])); WhatIf = $true } } `
    -Check { param($exit, $pt, $log) Assert-LogLacks $log 'Authenticating' 'WhatIf must not sign in'; "validated $(Split-Path -Leaf $state['zip'])" }

Invoke-Step -Name 'odata-pull' -Title 'Invoke-ProjectExport.ps1 -Mode OData (local template)' -Online -Script 'Invoke-ProjectExport.ps1' -Requires 'templateFolder' `
    -Params ($common + @{ LegalEntityId = $LegalEntityId; TemplateName = $TemplateName; ResourcesPath = $dirs.resources; Mode = 'OData'; DataPath = $dirs.data; Force = $true }) `
    -Check {
        param($exit, $pt, $log)
        $pulled = @($pt | Where-Object Status -in 'Pulled', 'Truncated')
        if ($pulled.Count -eq 0) { throw "nothing pulled: " + (($pt | Group-Object Status | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ' ') }
        $folder = Split-Path -Parent $pulled[0].File
        if (-not (Test-Path (Join-Path $folder '_pull.json'))) { throw '_pull.json missing' }
        Assert-LogLacks $log 'HTTP 501 - retrying' 'a 501 must not be retried'
        $state['dataFolder1'] = $folder
        (($pt | Group-Object Status | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ' ') + " -> $folder"
    }

Invoke-Step -Name 'odata-pull-env' -Title 'Invoke-ProjectExport.ps1 -Mode OData -TemplateSource Environment' -Online -Script 'Invoke-ProjectExport.ps1' `
    -Params ($common + @{ LegalEntityId = $LegalEntityId; TemplateName = $TemplateName; TemplateSource = 'Environment'; ResourcesPath = $dirs.resources; Mode = 'OData'; DataPath = $dirs.data; Force = $true }) `
    -Check {
        param($exit, $pt, $log)
        $pulled = @($pt | Where-Object Status -in 'Pulled', 'Truncated')
        if ($pulled.Count -eq 0) { throw 'nothing pulled' }
        if (-not $state.ContainsKey('dataFolder1')) { $state['dataFolder1'] = Split-Path -Parent $pulled[0].File }
        (($pt | Group-Object Status | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ' ')
    }

if ($SecondLegalEntityId) {
    Invoke-Step -Name 'odata-pull-2' -Title "Invoke-ProjectExport.ps1 -Mode OData for $SecondLegalEntityId" -Online -Script 'Invoke-ProjectExport.ps1' -Requires 'templateFolder' `
        -Params ($common + @{ LegalEntityId = $SecondLegalEntityId; TemplateName = $TemplateName; ResourcesPath = $dirs.resources; Mode = 'OData'; DataPath = $dirs.data; Force = $true }) `
        -Check {
            param($exit, $pt, $log)
            $pulled = @($pt | Where-Object Status -in 'Pulled', 'Truncated')
            if ($pulled.Count -eq 0) { throw 'nothing pulled' }
            $state['dataFolder2'] = Split-Path -Parent $pulled[0].File
            "$($pulled.Count) pulled -> $($state['dataFolder2'])"
        }
}

Invoke-Step -Name 'compare' -Title 'Compare-EnvironmentData.ps1 on the pulled folders' -Requires 'dataFolder1' -NoPassThru -Body {
    $f1 = $state['dataFolder1']
    $html = Join-Path $dirs.reports 'compare-self.html'
    & $hostExe -NoProfile -ExecutionPolicy Bypass -Command "& '$(Join-Path $repo 'Compare-EnvironmentData.ps1')' -Folder1 '$f1' -Folder2 '$f1' -HtmlPath '$html' -LogPath '' -FailOnDifference; exit `$LASTEXITCODE" *> (Join-Path $dirs.logs 'compare.self.log')
    if ($LASTEXITCODE -ne 0) { throw "self-compare reported differences (exit $LASTEXITCODE)" }
    $detail = 'self-compare identical'
    if ($state.ContainsKey('dataFolder2')) {
        $f2 = $state['dataFolder2']
        $html2 = Join-Path $dirs.reports 'compare-companies.html'; $json2 = Join-Path $dirs.reports 'compare-companies.json'; $xml = Join-Path $dirs.logs 'compare.passthru.xml'
        & $hostExe -NoProfile -ExecutionPolicy Bypass -Command "& '$(Join-Path $repo 'Compare-EnvironmentData.ps1')' -Folder1 '$f1' -Folder2 '$f2' -IgnoreFields dataAreaId -Template '$($TemplateName.Replace("'", "''"))' -HtmlPath '$html2' -JsonPath '$json2' -LogPath '' -PassThru | Export-Clixml '$xml' -Depth 6; exit `$LASTEXITCODE" *> (Join-Path $dirs.logs 'compare.companies.log')
        if ($LASTEXITCODE -ne 0) { throw "cross-company compare exit $LASTEXITCODE" }
        $f = @(if (Test-Path $xml) { Import-Clixml $xml })
        $summary = if ($f.Count -eq 0) { 'no differences (expected for a template of shared setup entities)' } else { ($f | Group-Object ChangeType | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ' ' }
        $detail += "; $LegalEntityId vs ${SecondLegalEntityId}: $summary -> $html2"
    }
    $detail
} -Check { param($exit, $pt, $log) $pt[-1] }

if ($TestDeviceCode) {
    $deviceCodeCommon = $common.Clone(); $deviceCodeCommon['AuthMode'] = 'DeviceCode'
    Invoke-Step -Name 'auth-devicecode' -Title 'one sign-in with -AuthMode DeviceCode (type the code when prompted)' -Online -Script 'Invoke-EnvironmentProbe.ps1' `
        -Params ($deviceCodeCommon + @{ LegalEntityId = $LegalEntityId; OutputPath = '' }) `
        -Check { param($exit, $pt, $log) Assert-LogHas $log 'via DeviceCode' 'device-code mode'; 'signed in via DeviceCode' }
}

if ($AllowImport) {
    $importCommon = @{ EnvironmentUrl = $ImportEnvironmentUrl; TenantId = $TenantId; AuthMode = $AuthMode; LogPath = '' }
    Invoke-Step -Name 'import' -Title "Invoke-BaselineImport.ps1 into $ImportEnvironmentUrl / $ImportLegalEntityId" -Online -Script 'Invoke-BaselineImport.ps1' -Requires 'expandedFolder' `
        -Params { $importCommon + @{ LegalEntityId = $ImportLegalEntityId; ResourcesPath = $dirs.packages; PackageName = (Split-Path -Leaf ([string]$state['expandedFolder'])); PollIntervalSeconds = $PollIntervalSeconds; Force = $true } } `
        -Check { param($exit, $pt, $log) $r = $pt[0]; if ($r.Status -notin 'Succeeded', 'PartiallySucceeded') { throw "status $($r.Status) (execution $($r.ExecutionId))" }; $state['imported'] = $true; "$($r.Status) execution $($r.ExecutionId)" }

    Invoke-Step -Name 'upload' -Title "Invoke-PackageUpload.ps1 into $ImportEnvironmentUrl / $ImportLegalEntityId" -Online -Script 'Invoke-PackageUpload.ps1' -Requires 'zip' `
        -Params { $importCommon + @{ LegalEntityId = $ImportLegalEntityId; UploadPath = $dirs.downloads; PackageName = (Split-Path -Leaf ([string]$state['zip'])); PollIntervalSeconds = $PollIntervalSeconds; Force = $true } } `
        -Check { param($exit, $pt, $log) $r = $pt[0]; if ($r.Status -notin 'Succeeded', 'PartiallySucceeded') { throw "status $($r.Status)" }; "$($r.Status) execution $($r.ExecutionId)" }

    Invoke-Step -Name 'job-report' -Title 'Get-ExecutionJobReport.ps1 on the import environment' -Online -Script 'Get-ExecutionJobReport.ps1' -Requires 'imported' `
        -Params ($importCommon + @{ HtmlPath = (Join-Path $dirs.reports 'job-report.html') }) `
        -Check { param($exit, $pt, $log) if (-not (Test-Path (Join-Path $dirs.reports 'job-report.html'))) { throw 'HTML not written' }; "$(@($pt).Count) detail rows -> job-report.html" }
}

# =============================================================================
#  Summary + report
# =============================================================================
Write-Host ''
Write-Rule "Acceptance summary  --  $($results.Count) steps"
$fmt = '  {0,-18} {1,-8} {2,-9} {3}'
Write-Host ($fmt -f 'Step', 'Status', 'Elapsed', 'Detail') -ForegroundColor White
foreach ($r in $results) {
    $c = switch ($r.Status) { 'PASS' { 'Green' } 'FAIL' { 'Red' } 'PLANNED' { 'Cyan' } default { 'DarkGray' } }
    $d = if ($r.Detail.Length -gt ($Script:LineWidth - 40)) { $r.Detail.Substring(0, $Script:LineWidth - 43) + '...' } else { $r.Detail }
    Write-Host ($fmt -f $r.Step, $r.Status, $r.Elapsed, $d) -ForegroundColor $c
}
$nPass = @($results | Where-Object Status -eq 'PASS').Count
$nFail = @($results | Where-Object Status -eq 'FAIL').Count
$nSkip = @($results | Where-Object Status -in 'SKIP', 'PLANNED').Count
Write-Host ''
Write-Host "  PASS $nPass   FAIL $nFail   SKIP $nSkip" -ForegroundColor $(if ($nFail) { 'Red' } else { 'Green' })

if (-not $WhatIf) {
    $md = [System.Text.StringBuilder]::new()
    [void]$md.AppendLine("# Acceptance test $ts")
    [void]$md.AppendLine()
    $tick = '`'
    [void]$md.AppendLine('| | |')
    [void]$md.AppendLine('|---|---|')
    [void]$md.AppendLine("| Repo | $tick$repo$tick |")
    [void]$md.AppendLine("| Environment | $tick$EnvironmentUrl$tick |")
    [void]$md.AppendLine("| Company | $LegalEntityId$(if ($SecondLegalEntityId) { ', ' + $SecondLegalEntityId }) |")
    [void]$md.AppendLine("| Template | $TemplateName |")
    [void]$md.AppendLine("| Host | $hostExe $($PSVersionTable.PSVersion) |")
    [void]$md.AppendLine("| Work path | $tick$WorkPath$tick |")
    [void]$md.AppendLine("| Result | PASS $nPass / FAIL $nFail / SKIP $nSkip |")
    [void]$md.AppendLine()
    [void]$md.AppendLine('| Step | Status | Elapsed | Detail | Log |')
    [void]$md.AppendLine('|---|---|---|---|---|')
    foreach ($r in $results) {
        $logCell    = if ($r.Log -and (Test-Path $r.Log)) { '`' + (Split-Path -Leaf $r.Log) + '`' } else { '' }
        $detailCell = $r.Detail -replace '\|', '\|'
        [void]$md.AppendLine("| $($r.Step) | **$($r.Status)** | $($r.Elapsed) | $detailCell | $logCell |")
    }
    [void]$md.AppendLine()
    [void]$md.AppendLine('## Commands')
    foreach ($r in $results | Where-Object { $_.Command -and $_.Command -ne '<in-process>' }) { [void]$md.AppendLine(); [void]$md.AppendLine("**$($r.Step)**"); [void]$md.AppendLine(); [void]$md.AppendLine('```'); [void]$md.AppendLine($r.Command); [void]$md.AppendLine('```') }
    $reportMd = Join-Path $dirs.reports 'AcceptanceReport.md'
    [System.IO.File]::WriteAllText($reportMd, $md.ToString(), (New-Object System.Text.UTF8Encoding($false)))
    $results | Select-Object Step, Title, Status, Detail, Elapsed, Log, Command | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $dirs.reports 'AcceptanceReport.json') -Encoding UTF8
    Write-Info "Report   : $reportMd"
    Write-Info "Logs     : $($dirs.logs)"
}

if ($PassThru) { $results | Write-Output }
if ($nFail -gt 0) { exit 1 }
