<#
.SYNOPSIS
    Shared HTML helpers so every report in this toolset looks the same.

.DESCRIPTION
    Dot-source this file to import:

        ConvertTo-HtmlEncoded   -- escape text for HTML
        Get-DmfReportCss        -- the stylesheet used by Get-ExecutionJobReport.ps1,
                                   plus classes for diff tables
        New-DmfHtmlDocument     -- wrap header / body / footer in a complete page

    Reports are self-contained: no external stylesheets, scripts, or fonts,
    so a file opened from a network share renders fully offline.
#>

function ConvertTo-HtmlEncoded {
    param([AllowNull()][AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    return ($Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;')
}


function Get-DmfReportCss {
    return @'
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;font-size:13px;background:#f3f2f1;color:#323130;line-height:1.5}

/* Header */
.page-header{background:#1b3a6b;color:#fff;padding:22px 40px}
.page-header h1{font-size:20px;font-weight:600;margin-bottom:6px}
.page-header .meta{font-size:12px;opacity:.85}
.page-header .meta span{margin-right:28px;white-space:nowrap}
.page-header .sides{display:flex;gap:40px;margin-top:10px;flex-wrap:wrap}
.page-header .side{font-size:12px;background:rgba(255,255,255,.08);padding:8px 14px;border-radius:3px;min-width:280px}
.page-header .side b{display:block;font-size:11px;text-transform:uppercase;letter-spacing:.5px;opacity:.75;margin-bottom:2px}

/* Summary cards */
.summary{display:flex;gap:14px;padding:22px 40px 8px;flex-wrap:wrap}
.card{background:#fff;border-radius:3px;padding:14px 22px 12px;min-width:110px;box-shadow:0 1px 3px rgba(0,0,0,.09);border-top:3px solid #0078d4}
.card.ok   {border-top-color:#107c10}
.card.warn {border-top-color:#ca5010}
.card.err  {border-top-color:#d13438}
.card.add  {border-top-color:#107c10}
.card.rem  {border-top-color:#d13438}
.card.chg  {border-top-color:#c19c00}
.card .val {font-size:30px;font-weight:700;line-height:1.1}
.card .lbl {font-size:11px;color:#797775;text-transform:uppercase;letter-spacing:.5px;margin-top:4px}
.card.ok   .val{color:#107c10}
.card.warn .val{color:#ca5010}
.card.err  .val{color:#d13438}
.card.add  .val{color:#107c10}
.card.rem  .val{color:#d13438}
.card.chg  .val{color:#8a6d00}

/* Notices, sections */
.notice{margin:8px 40px 0;padding:10px 14px;background:#fff4ce;border-left:4px solid #ca5010;border-radius:3px;font-size:12px}
.notice.info{background:#eef4fb;border-left-color:#0078d4}
.section{padding:18px 40px 0}
.section h2{font-size:15px;font-weight:600;color:#1b3a6b;margin-bottom:8px}
.section .sub{font-size:12px;color:#605e5c;margin-bottom:8px}
.toolbar{display:flex;gap:12px;align-items:center;padding:14px 40px 0;font-size:12px}
.toolbar input{padding:6px 10px;border:1px solid #c8c6c4;border-radius:3px;font-size:13px;min-width:320px}
.toolbar label{color:#605e5c}

/* Tables */
.table-wrap{background:#fff;border-radius:3px;box-shadow:0 1px 3px rgba(0,0,0,.09);overflow-x:auto}
table{border-collapse:collapse;width:100%;font-size:12.5px}
thead th{background:#e1dfdd;color:#323130;font-weight:600;text-align:left;padding:8px 12px;white-space:nowrap;position:sticky;top:0;z-index:1}
td{padding:7px 12px;border-bottom:1px solid #f0efee;vertical-align:top}
td.r{text-align:right;font-variant-numeric:tabular-nums}
td.mono,.mono{font-family:Consolas,'Cascadia Mono',monospace;font-size:12px}
tr:last-child td{border-bottom:none}
tr.r-ok   td{background:#fff}
tr.r-warn td{background:#fffdf5}
tr.r-err  td{background:#fff8f8}
tr.r-mute td{color:#8a8886}
tr.k-add  td{background:#f2fbf1}
tr.k-rem  td{background:#fdf3f3}
tr.k-chg  td{background:#fffbea}
.diff-old{color:#a4262c;text-decoration:line-through;text-decoration-color:rgba(164,38,44,.5)}
.diff-new{color:#107c10;font-weight:600}
.arrow{color:#a19f9d;padding:0 6px}
.kv{color:#605e5c}
.kv b{color:#323130;font-weight:600}

/* Entity blocks */
details.entity{background:#fff;border-radius:3px;box-shadow:0 1px 3px rgba(0,0,0,.09);margin:0 40px 12px}
details.entity summary{cursor:pointer;padding:10px 14px;font-weight:600;display:flex;gap:14px;align-items:center;list-style:none}
details.entity summary::-webkit-details-marker{display:none}
details.entity summary .name{flex:1}
details.entity summary .counts{font-weight:400;color:#605e5c;font-size:12px}
details.entity[open] summary{border-bottom:1px solid #edebe9}
details.entity .body{padding:0 0 6px}
details.entity .more{padding:8px 14px;color:#797775;font-style:italic;font-size:12px}

/* Badges */
.badge{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;font-weight:600;white-space:nowrap}
.b-ok  {background:#dff6dd;color:#107c10}
.b-err {background:#fde7e9;color:#c50f1f}
.b-warn{background:#fff4ce;color:#7a4f00}
.b-none{background:#f3f2f1;color:#605e5c}
.b-add {background:#dff6dd;color:#107c10}
.b-rem {background:#fde7e9;color:#c50f1f}
.b-chg {background:#fff4ce;color:#7a4f00}

/* Footer */
footer{padding:14px 40px;font-size:11px;color:#a19f9d;border-top:1px solid #edebe9;background:#faf9f8;margin-top:20px}

@media print{
  body{background:#fff}
  .page-header,thead th{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  .table-wrap,details.entity{box-shadow:none;border:1px solid #edebe9}
  thead th{position:static}
  details.entity{page-break-inside:avoid}
  .toolbar{display:none}
}
'@
}


function New-DmfHtmlDocument {
    <#
    .SYNOPSIS  Wraps header, body and footer fragments in a complete, self-contained page.
    #>
    param(
        [Parameter(Mandatory)][string]$Title,
        [AllowEmptyString()][string]$HeaderHtml = '',
        [AllowEmptyString()][string]$BodyHtml = '',
        [AllowEmptyString()][string]$FooterHtml = '',
        [AllowEmptyString()][string]$ExtraCss = '',
        [AllowEmptyString()][string]$ScriptJs = ''
    )
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html lang="en">')
    [void]$sb.AppendLine('<head>')
    [void]$sb.AppendLine('<meta charset="UTF-8">')
    [void]$sb.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1.0">')
    [void]$sb.AppendLine("<title>$(ConvertTo-HtmlEncoded $Title)</title>")
    [void]$sb.AppendLine('<style>')
    [void]$sb.AppendLine((Get-DmfReportCss))
    if ($ExtraCss) { [void]$sb.AppendLine($ExtraCss) }
    [void]$sb.AppendLine('</style>')
    [void]$sb.AppendLine('</head>')
    [void]$sb.AppendLine('<body>')
    if ($HeaderHtml) { [void]$sb.AppendLine($HeaderHtml) }
    [void]$sb.AppendLine($BodyHtml)
    if ($FooterHtml) { [void]$sb.AppendLine("<footer>$FooterHtml</footer>") }
    if ($ScriptJs)   { [void]$sb.AppendLine('<script>'); [void]$sb.AppendLine($ScriptJs); [void]$sb.AppendLine('</script>') }
    [void]$sb.AppendLine('</body>')
    [void]$sb.AppendLine('</html>')
    return $sb.ToString()
}
