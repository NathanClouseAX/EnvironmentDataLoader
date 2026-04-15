<#
.SYNOPSIS
    Console output and transcript helpers for D365 F&O DMF tooling.

.DESCRIPTION
    Dot-source this file to import all output formatting functions and the
    transcript management helper into your script.

    The following $Script: variables must be set by the caller before these
    functions are used:
        $Script:Version          -- displayed in Write-Banner
        $Script:LineWidth        -- controls rule/banner width (default 80)
        $Script:TranscriptActive -- managed by Stop-RunTranscript

.NOTES
    All Write-* functions write directly to the host (not the pipeline) so
    they are safe to call inside functions that return values.
#>

function Write-Banner {
    param(
        [switch]$DryRun,
        [string]$Title = 'D365 F&O Data Import Utility'
    )
    $w   = $Script:LineWidth
    $tag = if ($DryRun) { '  [WhatIf]' } else { '' }
    Write-Host ('=' * $w)                                 -ForegroundColor DarkCyan
    Write-Host "  $Title  v$Script:Version$tag"           -ForegroundColor Cyan
    Write-Host ('=' * $w)                                 -ForegroundColor DarkCyan
}

function Write-Rule {
    param([string]$Label = '')
    $w = $Script:LineWidth
    if ($Label) {
        $pad = [Math]::Max(1, $w - $Label.Length - 5)
        Write-Host "--- $Label $('-' * $pad)" -ForegroundColor DarkGray
    } else {
        Write-Host ('-' * $w) -ForegroundColor DarkGray
    }
}

function Write-Step   { param([string]$m) Write-Host "`n==> $m"    -ForegroundColor Cyan    }
function Write-Info   { param([string]$m) Write-Host "    $m"      -ForegroundColor Gray    }
function Write-Detail { param([string]$m) Write-Host "      $m"    -ForegroundColor DarkGray }
function Write-OK     { param([string]$m) Write-Host "`n[OK]   $m" -ForegroundColor Green   }
function Write-Warn   { param([string]$m) Write-Host "[WARN] $m"   -ForegroundColor Yellow  }
function Write-Fail   { param([string]$m) Write-Host "[FAIL] $m"   -ForegroundColor Red     }

function Format-Elapsed {
    <#
    .SYNOPSIS  Formats a TimeSpan as a compact human-readable string.
    .OUTPUTS   [string]  e.g. "2h 5m 30s", "12m 4s", "45s", "<1s"
    #>
    param([Parameter(Mandatory)][TimeSpan]$Elapsed)
    $h = [int][Math]::Floor($Elapsed.TotalHours)
    $m = $Elapsed.Minutes
    $s = $Elapsed.Seconds
    if ($h -gt 0) { return "${h}h ${m}m ${s}s" }
    if ($m -gt 0) { return "${m}m ${s}s" }
    if ($s -gt 0) { return "${s}s" }
    return '<1s'
}

function Stop-RunTranscript {
    <#
    .SYNOPSIS  Stops the PowerShell transcript if one was started by the caller.
    .NOTES     Reads and clears $Script:TranscriptActive in the calling script scope.
    #>
    if ($Script:TranscriptActive) {
        try { Stop-Transcript | Out-Null } catch {}
        $Script:TranscriptActive = $false
    }
}
