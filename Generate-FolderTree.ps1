<#
.SYNOPSIS
    Generate-FolderTree.ps1 v2 — PS 5.1 compatible
.IMPORTANT
    SAVE AS UTF-8 WITH BOM! Notepad++: Encoding -> Convert to UTF-8-BOM
#>
param(
    [string]$Path            = "",
    [string]$OutputFile      = "",
    [int]$MaxDepth           = 0,
    [switch]$IncludeHash,
    [string]$HashAlgorithm   = "SHA256",
    [int]$HashSizeLimitMB    = 200,
    [switch]$Update,
    [string[]]$DeletePattern,
    [string[]]$AddNote,
    [switch]$RestoreAll,
    [string[]]$RestorePattern,
    [switch]$Status
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
 $ErrorActionPreference = 'SilentlyContinue'

function E($t) {
    if ($null -eq $t) { return "" }
    $s = $t.ToString()
    $s = $s -replace '&','&amp;'
    $s = $s -replace '<','&lt;'
    $s = $s -replace '>','&gt;'
    $s = $s -replace '"','&quot;'
    $s = $s -replace "'",'&#39;'
    return $s
}

function FmtSz($b) {
    if ($null -eq $b -or $b -lt 0) { return "0 B" }
    if ($b -ge 1TB) { return ("{0:N2} " -f ($b/1TB)) + "TB" }
    if ($b -ge 1GB) { return ("{0:N2} " -f ($b/1GB)) + "GB" }
    if ($b -ge 1MB) { return ("{0:N2} " -f ($b/1MB)) + "MB" }
    if ($b -ge 1KB) { return ("{0:N2} " -f ($b/1KB)) + "KB" }
    return "$b B"
}

function SzClr($b, $d) {
    if ($d) {
        if ($b -ge 10GB) { return "sr" }
        if ($b -ge 1GB)  { return "so" }
        if ($b -ge 100MB){ return "sy" }
    } else {
        if ($b -ge 1GB)  { return "sr" }
        if ($b -ge 100MB){ return "so" }
        if ($b -ge 10MB) { return "sy" }
    }
    return ""
}

function SafeHash($fp, $al, $lm) {
    if ($lm -gt 0) {
        $sz = (Get-Item -LiteralPath $fp -EA 0).Length
        if ($null -ne $sz -and $sz -gt ($lm * 1MB)) {
            return "skip > $lm MB"
        }
    }
    try { return (Get-FileHash -LiteralPath $fp -Algorithm $al -EA Stop).Hash }
    catch { return "error" }
}

function ParseMeta($html) {
    if ($html -match '<!--FTMETA:(.+?):ENDFTMETA-->') {
        try { return $matches[1] | ConvertFrom-Json } catch { return $null }
    }
    return $null
}

function SerMeta($m) {
    $j = ($m | ConvertTo-Json -Depth 10 -Compress) -replace "`r`n", " "
    return "<!--FTMETA:${j}:ENDFTMETA-->"
}

function DoScan($rp, $md, $ih, $ha, $hlm) {
    $rp = [IO.Path]::GetFullPath($rp)
    $files = @{}
    $tsz = 0; $fc = 0; $dc = 0; $fsz = @{}
    $p = @{ LiteralPath = $rp; Force = $true; EA = 'SilentlyContinue' }
    if ($md -eq 0 -or $md -ge 1) {
        $p['Recurse'] = $true
        if ($md -gt 0) { $p['Depth'] = $md - 1 }
    }
    $items = Get-ChildItem @p | Where-Object { -not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) }
    foreach ($it in ($items | Where-Object { -not $_.PSIsContainer })) {
        $rel = $it.FullName.Substring($rp.Length).TrimStart('\','/') -replace '/','\'
        $sz = $it.Length
        $h = ""
        if ($ih) { $h = SafeHash $it.FullName $ha $hlm }
        $files[$rel] = @{ s=$sz; lw=$it.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"); h=$h; d=$false }
        $tsz += $sz; $fc++
        $parts = $rel -split '\\'; $cp = ""
        for ($i = 0; $i -lt $parts.Count - 1; $i++) {
            $cp += $parts[$i]
            if (-not $fsz.ContainsKey($cp)) { $fsz[$cp] = 0 }
            $fsz[$cp] += $sz
            $cp += '\'
        }
    }
    foreach ($it in ($items | Where-Object { $_.PSIsContainer }) | Sort-Object FullName) {
        $rel = $it.FullName.Substring($rp.Length).TrimStart('\','/') -replace '/','\'
        $ds = 0
        if ($fsz.ContainsKey($rel)) { $ds = $fsz[$rel] }
        $files[$rel] = @{ s=$ds; lw=$it.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"); h=""; d=$true }
        $dc++
    }
    return @{ Files=$files; TS=$tsz; FC=$fc; DC=$dc }
}

function MergeM($om, $ns) {
    $of = $om.f; $nf = $ns.Files
    $dp = @($om.dp)
    $prevStates = @{}
    if ($om.ps) {
        if ($om.ps -is [System.Collections.IDictionary]) {
            $prevStates = $om.ps
        } else {
            foreach ($p in $om.ps.PSObject.Properties) { $prevStates[$p.Name] = $p.Value }
        }
    }
    $st = @{}; $ac = 0; $rc = 0; $cc = 0
    $oldK = @($of.PSObject.Properties.Name)
    foreach ($k in $nf.Keys) {
        if ($dp -contains $k) { continue }
        if ($k -in $oldK) {
            $oe = $of.$k; $ne = $nf[$k]
            if ($prevStates.ContainsKey($k) -and $prevStates[$k] -eq "Gone") { $st[$k] = "New"; $ac++ }
            elseif ($oe.s -ne $ne.s -or $oe.lw -ne $ne.lw) { $st[$k] = "Changed"; $cc++ }
        } else { $st[$k] = "New"; $ac++ }
    }
    foreach ($k in $oldK) {
        if ($dp -contains $k) { continue }
        if (-not $nf.ContainsKey($k)) { $st[$k] = "Gone"; $rc++ }
    }
    $he = @{ d=(Get-Date -Format "yyyy-MM-dd HH:mm:ss"); a=$ac; r=$rc; c=$cc }
    if (-not $om.h) { $om.h = @() }
    $om.h += $he
    $mf = @{}
    foreach ($k in $oldK) {
        if ($dp -contains $k) { continue }
        if (-not $nf.ContainsKey($k)) {
            $v = $of.$k; $hv = ""
            if ($v.h) { $hv = $v.h }
            $mf[$k] = @{ s=$v.s; lw=$v.lw; h=$hv; d=$v.d }
        }
    }
    foreach ($k in $nf.Keys) {
        if ($dp -contains $k) { continue }
        $mf[$k] = $nf[$k]
    }
    $mfObj = [PSCustomObject]@{}
    foreach ($k in $mf.Keys) {
        $v = $mf[$k]
        $mfObj | Add-Member -NotePropertyName $k -NotePropertyValue ([PSCustomObject]@{ s=$v.s; lw=$v.lw; h=$v.h; d=$v.d })
    }
    $om.f = $mfObj
    $om | Add-Member -NotePropertyName 'ps' -NotePropertyValue $st -Force
    return @{ Files=$mf; States=$st; Meta=$om; Stats=@{ TS=$ns.TS; FC=$ns.FC; DC=$ns.DC; N=$ac; G=$rc; C=$cc; X=$dp.Count } }
}

function BuildHTML($meta, $states, $stats, $scriptPath, $outFile) {
    $dp = @($meta.dp); $notes = $meta.a
    $activeFiles = @{}; $exclFiles = @{}
    foreach ($prop in $meta.f.PSObject.Properties) {
        $k = $prop.Name; $v = $prop.Value
        $hv = ""
        if ($v.h) { $hv = $v.h }
        $entry = @{ s=$v.s; lw=$v.lw; h=$hv; d=$v.d }
        if ($dp -contains $k) { $exclFiles[$k] = $entry } else { $activeFiles[$k] = $entry }
    }
    $tsz = 0; $fc = 0; $dc = 0
    foreach ($v in $activeFiles.Values) {
        if ($v.d) { $dc++ } else { $fc++; $tsz += $v.s }
    }
    $nC = 0; $gC = 0; $cC = 0; $xC = $dp.Count
    if ($stats) { $tsz = $stats.TS; $fc = $stats.FC; $dc = $stats.DC; $nC = $stats.N; $gC = $stats.G; $cC = $stats.C; $xC = $stats.X }

    # -- pre-build all conditional HTML (NO ternary!) --
    $statNew = ""
    if ($nC -gt 0) { $statNew = "<span style='color:#3fb950'>&#128195; New: <b>$nC</b></span>" }
    $statGone = ""
    if ($gC -gt 0) { $statGone = "<span style='color:#f85149'>&#10060; Gone: <b>$gC</b></span>" }
    $statChg = ""
    if ($cC -gt 0) { $statChg = "<span style='color:#d29922'>&#9999; Changed: <b>$cC</b></span>" }
    $statExcl = ""
    if ($xC -gt 0) { $statExcl = "<span style='color:#6e7681'>&#128465; Excluded: <b>$xC</b></span>" }

    $hashInfo = ""
    if ($meta.p.ih) { $hashInfo = "<span>Hash: $($meta.p.ha)</span>" }
    $depthInfo = ""
    if ($meta.p.md -gt 0) { $depthInfo = "<span>Depth: $($meta.p.md)</span>" }

    $scanDate = ""
    if ($meta.sd) { $scanDate = $meta.sd } else { $scanDate = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
    $metaStr = ""

    $gnHtml = ""
    if ($notes -and $notes.Count -gt 0) {
        $hasG = $false
        foreach ($n in $notes) {
            if (-not $n.p) {
                if (-not $hasG) { $gnHtml += "<div class='notes-sec'><h2>&#128203; Notes</h2>`n"; $hasG = $true }
                $gnHtml += "<div class='unote gnote'><span class='ni'>&#128172;</span><span class='nt'>$(E $n.t)</span><span class='nd'>$(E $n.dt)</span></div>`n"
            }
        }
        if ($hasG) { $gnHtml += "</div>`n" }
    }

    $oldCustom = @()
    $oldDataJsPath = [IO.Path]::Combine([IO.Path]::GetDirectoryName($outFile), [IO.Path]::GetFileNameWithoutExtension($outFile) + ".data.js")
    if (Test-Path $oldDataJsPath) {
        $oldDataJsRaw = Get-Content $oldDataJsPath -Raw -Encoding UTF8
        if ($oldDataJsRaw -match 'var\s+TREE_DATA\s*=\s*') {
            $eqIdx = $oldDataJsRaw.IndexOf('=', $oldDataJsRaw.IndexOf('TREE_DATA')) + 1
            $jsonStr = $oldDataJsRaw.Substring($eqIdx).Trim()
            if ($jsonStr.EndsWith(';')) { $jsonStr = $jsonStr.Substring(0, $jsonStr.Length - 1).Trim() }
            try {
                $oldDataParsed = $jsonStr | ConvertFrom-Json
                if ($oldDataParsed.custom) { $oldCustom = @($oldDataParsed.custom) }
            } catch { }
        }
    }
    $fObj = [PSCustomObject]@{}
    foreach ($k in $activeFiles.Keys) {
        $v = $activeFiles[$k]
        $fObj | Add-Member -NotePropertyName $k -NotePropertyValue ([PSCustomObject]@{s=$v.s; lw=$v.lw; h=$v.h; d=$v.d})
    }
    $jsData = [PSCustomObject]@{
        sp  = $meta.sp
        scp = $meta.scp
        p   = $meta.p
        sd  = $scanDate
        f   = $fObj
        a   = @($meta.a)
        dp  = @($meta.dp)
        h   = @($meta.h)
        custom = $oldCustom
        states = @{}
        stats = [PSCustomObject]@{ts=$tsz; fc=$fc; dc=$dc; n=$nC; g=$gC; c=$cC; x=$xC}
    }
    if ($states) {
        $statesObj = [PSCustomObject]@{}
        foreach ($k in $states.Keys) {
            $statesObj | Add-Member -NotePropertyName $k -NotePropertyValue $states[$k]
        }
        $jsData.states = $statesObj
    }
    $treeJson = "{}"
	$dataJsName = [IO.Path]::GetFileNameWithoutExtension($outFile) + ".data.js"
    $dataJsPath = [IO.Path]::Combine([IO.Path]::GetDirectoryName($outFile), $dataJsName)
    if (Test-Path -LiteralPath $dataJsPath) {
        Copy-Item -LiteralPath $dataJsPath -Destination ($dataJsPath + ".bak") -Force
    }
    $dataJsContent = "var TREE_DATA = " + ($jsData | ConvertTo-Json -Depth 10) + ";"
    [IO.File]::WriteAllText($dataJsPath, $dataJsContent, [Text.Encoding]::UTF8)
    $dataJsNameEsc = E $dataJsName

    $histHtml = ""
    if ($meta.h -and $meta.h.Count -gt 0) {
        $histHtml = "<div class='hist-sec'><h3>&#128220; Update history</h3><ul>" + "`n"
        foreach ($he in $meta.h) { $histHtml += "<li>$($he.d): +$($he.a) new, -$($he.r) gone, ~$($he.c) changed</li>" + "`n" }
        $histHtml += "</ul></div>" + "`n"
    }

    $filterBtns = ""
    if ($nC -gt 0 -or $gC -gt 0 -or $cC -gt 0) {
        $filterBtns = "<div class='filters'>" + "`n"
        $filterBtns += "<button onclick='filterByState(&quot;&quot;)' class='fbtn active' id='fAll'>All</button>" + "`n"
        if ($nC -gt 0) { $filterBtns += "<button onclick='filterByState(&quot;New&quot;)' class='fbtn' id='fNew'>New ($nC)</button>" + "`n" }
        if ($cC -gt 0) { $filterBtns += "<button onclick='filterByState(&quot;Changed&quot;)' class='fbtn' id='fChg'>Changed ($cC)</button>" + "`n" }
        if ($gC -gt 0) { $filterBtns += "<button onclick='filterByState(&quot;Gone&quot;)' class='fbtn' id='fGon'>Gone ($gC)</button>" + "`n" }
        $filterBtns += "</div>" + "`n"
    }

    $spEsc = E $meta.sp
    $scriptEsc = E $scriptPath
    $ofEsc2 = E $outFile
    $szStr = FmtSz $tsz

    $cmdLine = "powershell -ExecutionPolicy Bypass -File " + '"' + $scriptEsc + '"' + " -Update -OutputFile " + '"' + $ofEsc2 + '"'

 $htmlBefore = @"
<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Folder Tree - $spEsc</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:#0d1117;color:#c9d1d9;padding:16px;max-width:1400px;margin:0 auto}
.header{background:#161b22;border:1px solid #30363d;border-radius:12px;padding:20px 24px;margin-bottom:16px}
.header h1{font-size:1.3em;color:#58a6ff;margin-bottom:8px;word-break:break-all}
.header h1 span{color:#8b949e;font-weight:400}
.stats{display:flex;flex-wrap:wrap;gap:12px 24px;margin:12px 0;font-size:.9em;color:#8b949e}
.stats b{color:#c9d1d9}
.scan-info{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin:12px 0;font-size:.85em;color:#8b949e}
.scan-info .date{color:#58a6ff}
.btn{background:#21262d;border:1px solid #30363d;color:#c9d1d9;padding:6px 14px;border-radius:6px;cursor:pointer;font-size:.85em;transition:all .15s}
.btn:hover{background:#30363d;border-color:#58a6ff;color:#58a6ff}
.btn-primary{background:#238636;border-color:#238636;color:#fff}
.btn-primary:hover{background:#2ea043}
.controls{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-top:12px}
#search{background:#0d1117;border:1px solid #30363d;color:#c9d1d9;padding:6px 12px;border-radius:6px;font-size:.9em;width:260px;outline:none}
#search:focus{border-color:#58a6ff}
#search::placeholder{color:#484f58}
.filters{display:flex;gap:6px;flex-wrap:wrap;margin-top:10px}
.fbtn{background:#0d1117;border:1px solid #30363d;color:#8b949e;padding:4px 10px;border-radius:4px;cursor:pointer;font-size:.8em}
.fbtn:hover,.fbtn.active{background:#1f6feb;border-color:#1f6feb;color:#fff}
.notes-sec{background:#161b22;border:1px solid #f0883e33;border-left:4px solid #f0883e;border-radius:8px;padding:16px 20px;margin-bottom:16px}
.notes-sec h2{font-size:1em;color:#f0883e;margin-bottom:10px}
.unote{background:#f0883e15;border:1px solid #f0883e30;border-radius:6px;padding:8px 12px;margin:6px 0;display:flex;align-items:flex-start;gap:8px;font-size:.9em}
.unote.gnote{background:#f0883e20;border-color:#f0883e50}
.ni{flex-shrink:0;font-size:1.1em}
.nt{flex:1;color:#e6c99a;line-height:1.4}
.nd{flex-shrink:0;color:#484f58;font-size:.8em;white-space:nowrap}
.tree{background:#161b22;border:1px solid #30363d;border-radius:12px;padding:12px 8px}
.de{margin:2px 0}
.de>summary{list-style:none;cursor:pointer;padding:4px 8px;border-radius:6px;display:flex;align-items:center;gap:6px;font-size:.9em;user-select:none}
.de>summary:hover{background:#1c2128}
.de>summary::-webkit-details-marker{display:none}
.de>summary::before{content:'\25B6';font-size:.65em;color:#484f58;transition:transform .15s;display:inline-block;min-width:12px}
.de[open]>summary::before{transform:rotate(90deg)}
.fc{margin-left:20px;border-left:1px solid #21262d;padding-left:4px}
.fe{display:flex;align-items:center;gap:6px;padding:3px 8px;border-radius:4px;font-size:.88em;margin:1px 0}
.fe:hover{background:#1c2128}
.sb{width:4px;height:20px;border-radius:2px;flex-shrink:0}
.sy{background:#d29922}.so{background:#db6d28}.sr{background:#f85149}
.ic{flex-shrink:0;font-size:1em}
.nm{color:#c9d1d9;word-break:break-all;flex:1;min-width:0}
.mt{color:#484f58;font-size:.8em;white-space:nowrap;flex-shrink:0}
.badge{padding:1px 6px;border-radius:3px;font-size:.7em;font-weight:600;white-space:nowrap;flex-shrink:0}
.bn{background:#23863633;color:#3fb950;border:1px solid #23863655}
.bc{background:#d2992233;color:#d29922;border:1px solid #d2992255}
.bg{background:#f8514933;color:#f85149;border:1px solid #f8514955}
.ab{background:none;border:none;color:#6e7681;cursor:pointer;font-size:.85em;padding:2px 6px;border-radius:4px;flex-shrink:0;margin-left:2px}
.ab:hover{background:#f8514933;color:#f85149}
.ab.note:hover{background:#58a6ff33;color:#58a6ff}
.st-gone{opacity:.55}
.st-gone .nm{text-decoration:line-through;color:#6e7681}
.excl-sec{background:#161b22;border:1px solid #30363d;border-radius:12px;padding:12px 16px;margin-top:16px}
.excl-sum{cursor:pointer;color:#f85149;font-weight:600;font-size:.95em;padding:8px;list-style:none;display:flex;align-items:center;gap:8px}
.excl-sum::-webkit-details-marker{display:none}
.excl-sum::before{content:'\25B6';font-size:.65em;transition:transform .15s}
details[open]>.excl-sum::before{transform:rotate(90deg)}
.excl-body{margin-left:20px;opacity:.5;padding:8px 0}
.excl-body .fe .nm{text-decoration:line-through}
.excl-hint{font-size:.78em;color:#484f58;margin-top:8px;padding:8px;background:#0d1117;border-radius:6px;font-family:'Cascadia Code',Consolas,monospace;word-break:break-all}
.dl{color:#d29922;font-size:.85em;padding:6px 12px;font-style:italic}
.hist-sec{margin-top:16px;background:#161b22;border:1px solid #30363d;border-radius:12px;padding:16px 20px}
.hist-sec h3{font-size:.95em;color:#bc8cff;margin-bottom:8px}
.hist-sec ul{list-style:none;padding:0}
.hist-sec li{font-size:.85em;color:#8b949e;padding:3px 0;border-bottom:1px solid #21262d33}
.hist-sec li:last-child{border:none}
.rescan-sec{margin-top:16px;background:#161b22;border:1px solid #30363d;border-radius:12px;padding:16px 20px;text-align:center}
.rescan-sec h3{font-size:.95em;color:#58a6ff;margin-bottom:12px}
.rescan-cmd{background:#0d1117;border:1px solid #30363d;border-radius:6px;padding:10px 14px;font-family:'Cascadia Code',Consolas,monospace;font-size:.82em;color:#7ee787;word-break:break-all;text-align:left;margin:10px 0;position:relative;cursor:pointer}
.rescan-cmd .copy-hint{position:absolute;right:8px;top:50%;transform:translateY(-50%);color:#484f58;font-size:.75em;font-family:'Segoe UI',sans-serif}
.toast{position:fixed;bottom:20px;right:20px;background:#238636;color:#fff;padding:10px 18px;border-radius:8px;font-size:.9em;opacity:0;transition:opacity .3s;pointer-events:none;z-index:999}
.toast.show{opacity:1}
.hidden{display:none!important}
.controls2{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin:12px 0;padding:12px 16px;background:#161b22;border:1px solid #30363d;border-radius:8px}
.fe input[type=checkbox],.de>summary input[type=checkbox]{margin-right:4px;accent-color:#58a6ff;display:none}
.show-cb .fe input[type=checkbox],.show-cb .de>summary input[type=checkbox]{display:inline}
.custom-row .nm{color:#58a6ff;font-style:italic}
.custom-row .sb{background:#58a6ff}
.sel-count{color:#f85149;font-size:.85em;margin-left:8px}
.add-btn{background:none;border:none;color:#484f58;cursor:pointer;font-size:1.1em;padding:0 4px;border-radius:4px;flex-shrink:0;line-height:1}
.add-btn:hover{background:#23863633;color:#3fb950}
.modal-bg{position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.6);z-index:100;display:flex;align-items:center;justify-content:center}
.modal{background:#161b22;border:1px solid #30363d;border-radius:12px;padding:20px;min-width:360px;max-width:500px}
.modal h3{color:#58a6ff;margin-bottom:14px;font-size:1em}
.modal-row{display:flex;gap:8px;margin-bottom:8px}
.modal-row input{flex:1;background:#0d1117;border:1px solid #30363d;color:#c9d1d9;padding:8px 10px;border-radius:6px;font-size:.9em;font-family:inherit}
.modal-hint{font-size:.8em;color:#484f58;margin-bottom:10px}
</style>
</head>
<body>
 $metaStr
<div class="header">
<h1>Folder Tree <span>- $spEsc</span></h1>
<div class="stats">
<span>&#128193; Folders: <b>$dc</b></span>
<span>&#128196; Files: <b>$fc</b></span>
<span>&#128190; Total: <b>$szStr</b></span>
 $statNew$statGone$statChg$statExcl
</div>
<div class="scan-info">
<span>Scanned: <span class="date">$scanDate</span></span>
 $hashInfo$depthInfo
</div>
<div class="controls">
<input type="text" id="search" placeholder="Search..." oninput="doSearch()">
<button class="btn" onclick="expandAll()">Expand all</button>
<button class="btn" onclick="collapseAll()">Collapse all</button>
</div>
 $filterBtns
</div>
 $gnHtml
<div class="controls2" id="ctrlBar">
<button class="btn" onclick="showAddMenu('')">+ Строка в корень</button>
<button class="btn" onclick="addCustomFolder('')">+ Папка в корень</button>
<button class="btn" id="btnDel" onclick="deleteSelected()" style="display:none">Удалить выбранные <span class="sel-count" id="selCount"></span></button>
<button class="btn btn-primary" onclick="saveChanges()">💾 Сохранить</button>
<label style="font-size:.85em;color:#8b949e;margin-left:12px;cursor:pointer"><input type="checkbox" id="showCB" onchange="toggleCB()"> Чекбоксы</label>
</div>
<div class="tree" id="tree"></div>
<script src="$dataJsNameEsc" onerror="window._dataJsFailed=true"></script>
<script type="application/json" id="tree-data">$treeJson</script>
"@

    $htmlAfter = @"
</div>
 $histHtml
<div class="rescan-sec">
<h3>&#128260; Rescan and update</h3>
<p style="font-size:.85em;color:#8b949e;margin-bottom:8px">Скопируйте команду и вставьте в PowerShell:</p>
<div class="rescan-cmd" id="cmdBox" onclick="copyCmd()">$cmdLine<span class="copy-hint">click to copy</span></div>
</div>
<div class="toast" id="toast">Copied!</div>
<script>
var D,DATA_JS_NAME='$dataJsNameEsc',SCRIPT_PATH='$scriptPathJs',OUT_FILE='$outFileJs';
(function(){
if(typeof TREE_DATA!=='undefined'){D=TREE_DATA;}
else{try{D=JSON.parse(document.getElementById('tree-data').textContent);}catch(e){alert('JSON load error');return;}}
if(!D.custom)D.custom=[];
render();
})();

function buildTree(files){
var root={_n:'',_c:{},_e:null};
if(!files)return root;
for(var k in files){
if(!files.hasOwnProperty(k))continue;
var parts=k.split('\\'),cur=root;
for(var i=0;i<parts.length;i++){
var p=parts[i];
if(!cur._c[p])cur._c[p]={_n:p,_c:{},_e:null};
cur=cur._c[p];
}
cur._e=files[k];
}
return root;
}

function fmtSz(b){
if(b==null||b<0)return '0 B';
if(b>=1099511627776)return(b/1099511627776).toFixed(2)+' TB';
if(b>=1073741824)return(b/1073741824).toFixed(2)+' GB';
if(b>=1048576)return(b/1048576).toFixed(2)+' MB';
if(b>=1024)return(b/1024).toFixed(2)+' KB';
return b+' B';
}

function szClr(b,isDir){
if(isDir){if(b>=10737418240)return'sr';if(b>=1073741824)return'so';if(b>=104857600)return'sy';}
else{if(b>=1073741824)return'sr';if(b>=104857600)return'so';if(b>=10485760)return'sy';}
return'';
}

function esc(s){if(!s)return'';return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}
function mkBadge(s){
if(s==='New')return'<span class="badge bn">NEW</span>';
if(s==='Changed')return'<span class="badge bc">CHANGED</span>';
if(s==='Gone')return'<span class="badge bg">GONE</span>';
return'';
}
function getCustomChildren(parentId){
if(!D.custom||!D.custom.length)return[];
return D.custom.filter(function(c){return c.parent===parentId;});
}

function renderNode(node,depth,parentPath,maxD){
var h='';
var dirs=[],files=[];
for(var k in node._c){
if(!node._c.hasOwnProperty(k))continue;
var ch=node._c[k];
if(ch._e&&ch._e.d)dirs.push(k);else files.push(k);
}
dirs.sort();files.sort();
var all=dirs.concat(files);
for(var i=0;i<all.length;i++){
var k=all[i],ch=node._c[k],e=ch._e;
var fp=parentPath?parentPath+'\\'+k:k;
if(!e){
h+='<details class="de"><summary><span class="ic">📁</span><span class="nm">'+esc(k)+'</span></summary>';
h+=renderNode(ch,depth+1,fp,maxD);
h+='</details>\n';
continue;
}
if(e.d){
var st=D.states&&D.states[fp]?D.states[fp]:null;
var stCls=st?'st-'+st.toLowerCase():'';
h+='<details class="de '+stCls+'" data-fp="'+esc(fp)+'">\n<summary>';
h+='<input type="checkbox" data-fp="'+esc(fp)+'" onclick="event.stopPropagation();updateSelCount()">';
h+='<span class="sb '+szClr(e.s,true)+'"></span>';
h+='<span class="ic">📁</span>';
h+='<span class="nm">'+esc(k)+'</span>';
h+='<span class="mt">'+fmtSz(e.s)+'</span>';
if(st)h+=mkBadge(st);
h+='<button class="add-btn" title="Добавить..." onclick="event.stopPropagation();showAddMenu(\''+esc(fp).replace(/&#39;/g,"\\'")+'\')">&#10010;</button>';
h+='</summary>\n';
if(maxD===0||depth<maxD){
h+='<div class="fc">\n';
h+=renderNode(ch,depth+1,fp,maxD);
h+=renderCustom(fp);
h+='</div>\n';
}
h+='</details>\n';
}else{
var st2=D.states&&D.states[fp]?D.states[fp]:null;
var stCls2=st2?'st-'+st2.toLowerCase():'';
h+='<div class="fe '+stCls2+'" data-fp="'+esc(fp)+'">\n';
h+='<input type="checkbox" data-fp="'+esc(fp)+'" onclick="event.stopPropagation();updateSelCount()">';
h+='<span class="sb '+szClr(e.s,false)+'"></span>';
h+='<span class="ic">📄</span>';
h+='<span class="nm">'+esc(k)+'</span>';
h+='<span class="mt">'+fmtSz(e.s);
if(e.lw)h+=' | '+e.lw;
if(e.h)h+=' | '+e.h;
h+='</span>';
if(st2)h+=mkBadge(st2);
h+='</div>\n';
}
}
return h;
}

function renderCustom(parentId){
var items=getCustomChildren(parentId);
if(!items.length)return'';
var h='';
for(var i=0;i<items.length;i++){
var c=items[i];
if(c.type==='folder'){
h+='<details class="de custom-row" data-cid="'+c.id+'">\n<summary>';
h+='<span class="sb"></span>';
h+='<span class="ic">📁</span>';
h+='<span class="nm">'+esc(c.name)+'</span>';
h+='<button class="add-btn" title="Добавить..." onclick="event.stopPropagation();showAddMenu(\''+c.id+'\')">&#10010;</button>';
h+='<button class="add-btn" style="color:#f85149" title="Удалить" onclick="event.stopPropagation();deleteCustom(\''+c.id+'\')">&#128465;</button>';
h+='</summary>\n<div class="fc">\n';
h+=renderCustom(c.id);
h+='</div>\n</details>\n';
}else{
h+='<div class="fe custom-row" data-cid="'+c.id+'">\n';
h+='<span class="sb"></span><span class="ic">📝</span>';
h+='<span class="nm">'+esc(c.name)+'</span>';
h+='<span class="mt" style="color:#58a6ff">заметка</span>';
h+='<button class="add-btn" style="color:#f85149" title="Удалить" onclick="event.stopPropagation();deleteCustom(\''+c.id+'\')">&#128465;</button>';
h+='</div>\n';
}
}
return h;
}

function render(){
var tree=buildTree(D.f);
var maxD=D.p&&D.p.md?D.p.md:0;
var container=document.getElementById('tree');
container.innerHTML=renderNode(tree,0,'',maxD)+renderCustom('');
updateSelCount();
}

function toggleCB(){document.body.classList.toggle('show-cb',document.getElementById('showCB').checked);}

function getSelected(){
var cbs=document.querySelectorAll('.tree input[type=checkbox]:checked');
var fps=[];
for(var i=0;i<cbs.length;i++){var fp=cbs[i].getAttribute('data-fp');if(fp)fps.push(fp);}
return fps;
}

function updateSelCount(){
var n=getSelected().length;
var btn=document.getElementById('btnDel');
var cnt=document.getElementById('selCount');
if(n>0){btn.style.display='';cnt.textContent=' ('+n+')';}else{btn.style.display='none';}
}

function deleteSelected(){
var fps=getSelected();
if(!fps.length)return;
if(!confirm('Удалить '+fps.length+' записей?\n\n'+fps.slice(0,5).join('\n')+(fps.length>5?'\n...ещё '+(fps.length-5):'')))return;
for(var i=0;i<fps.length;i++){delete D.f[fps[i]];}
render();
showToast('Удалено: '+fps.length);
}

function showAddMenu(parentId){
var bg=document.createElement('div');
bg.className='modal-bg';
bg.onclick=function(ev){if(ev.target===bg)bg.remove();};
var label=parentId?parentId:'корень';
bg.innerHTML='<div class="modal">'+
'<h3>Добавить в: '+esc(label)+'</h3>'+
'<div class="modal-hint">Путь к папке или id виртуальной папки</div>'+
'<div class="modal-row"><input type="text" id="mInput" placeholder="Название..." autofocus></div>'+
'<div style="display:flex;gap:8px;justify-content:flex-end;margin-top:14px">'+
'<button class="btn" onclick="this.closest(\'.modal-bg\').remove()">Отмена</button>'+
'<button class="btn" onclick="doAddNote(\''+parentId.replace(/'/g,"\\'")+'\')">+ Строка</button>'+
'<button class="btn" onclick="doAddFolder(\''+parentId.replace(/'/g,"\\'")+'\')">+ Папка</button>'+
'</div></div>';
document.body.appendChild(bg);
var inp=bg.querySelector('#mInput');
inp.focus();
inp.onkeydown=function(ev){if(ev.key==='Enter')doAddNote(parentId);};
}

function doAddNote(parentId){
var inp=document.querySelector('.modal #mInput');
var text=inp?inp.value.trim():'';
if(!text){inp.focus();return;}
D.custom.push({id:'c'+Date.now(),parent:parentId,name:text,type:'note'});
document.querySelector('.modal-bg').remove();
render();
showToast('Строка добавлена');
}

function doAddFolder(parentId){
var inp=document.querySelector('.modal #mInput');
var name=inp?inp.value.trim():'';
if(!name){inp.focus();return;}
D.custom.push({id:'c'+Date.now(),parent:parentId,name:name,type:'folder'});
document.querySelector('.modal-bg').remove();
render();
showToast('Папка добавлена');
}

function addCustomFolder(parentId){
var name=prompt('Имя новой папки:');
if(!name)return;
D.custom.push({id:'c'+Date.now(),parent:parentId,name:name,type:'folder'});
render();
showToast('Папка добавлена');
}

function deleteCustom(id){
if(!confirm('Удалить эту запись и всё внутри?'))return;
function collectIds(pid){
var ids=[pid];
for(var i=0;i<D.custom.length;i++){
if(D.custom[i].parent===pid)ids=ids.concat(collectIds(D.custom[i].id));
}
return ids;
}
var rm=collectIds(id);
D.custom=D.custom.filter(function(c){return rm.indexOf(c.id)===-1;});
render();
showToast('Удалено');
}

function saveChanges(){
var js='var TREE_DATA = '+JSON.stringify(D,null,2)+';';
var blob=new Blob([js],{type:'application/javascript;charset=utf-8'});
var a=document.createElement('a');
a.href=URL.createObjectURL(blob);
a.download=DATA_JS_NAME;
document.body.appendChild(a);a.click();document.body.removeChild(a);
URL.revokeObjectURL(a.href);
showToast(DATA_JS_NAME+' сохранён — замените старый файл');
}

function doSearch(){var q=document.getElementById('search').value.toLowerCase();var all=document.querySelectorAll('.tree .fe,.tree .de');if(!q){for(var i=0;i<all.length;i++){all[i].classList.remove('hidden');all[i].style.display='';}var ds=document.querySelectorAll('.tree .de');for(var i=0;i<ds.length;i++){ds[i].removeAttribute('data-search-hidden');}return;}for(var i=0;i<all.length;i++){all[i].classList.remove('hidden');all[i].style.display='';}for(var i=0;i<all.length;i++){var nm=all[i].querySelector('.nm');if(!nm)continue;if(nm.textContent.toLowerCase().indexOf(q)===-1){all[i].style.display='none';all[i].classList.add('hidden');}}var ds=document.querySelectorAll('.tree .de');for(var i=0;i<ds.length;i++){var d=ds[i];var vis=d.querySelectorAll('.fe:not(.hidden),.de:not(.hidden):not([data-search-hidden])');if(vis.length===0){d.style.display='none';d.setAttribute('data-search-hidden','1');}}}
function expandAll(){var ds=document.querySelectorAll('.tree .de');for(var i=0;i<ds.length;i++)ds[i].open=true;}
function collapseAll(){var ds=document.querySelectorAll('.tree .de');for(var i=0;i<ds.length;i++)ds[i].open=false;}
function filterByState(s){var btns=document.querySelectorAll('.fbtn');for(var i=0;i<btns.length;i++)btns[i].classList.remove('active');if(s){var id='f'+s.charAt(0).toUpperCase()+s.slice(1);var b=document.getElementById(id);if(b)b.classList.add('active');}else{var ba=document.getElementById('fAll');if(ba)ba.classList.add('active');}var all=document.querySelectorAll('.tree .fe,.tree .de');for(var i=0;i<all.length;i++){var e=all[i];if(!s){e.classList.remove('hidden');e.style.display='';continue;}if(e.classList.contains('st-'+s.toLowerCase())){e.classList.remove('hidden');e.style.display='';}else{e.style.display='none';e.classList.add('hidden');}}if(s){var ds=document.querySelectorAll('.tree .de');for(var i=0;i<ds.length;i++){var d=ds[i];var vis=d.querySelectorAll('.fe:not(.hidden),.de:not(.hidden)');d.style.display=vis.length===0?'none':'';}}else{var ds2=document.querySelectorAll('.tree .de');for(var i=0;i<ds2.length;i++)ds2[i].style.display='';}}
function copyCmd(){var box=document.getElementById('cmdBox');var t=box.innerText.replace('click to copy','').trim();navigator.clipboard.writeText(t).then(function(){showToast('Скопировано');});}
function showToast(msg){var to=document.getElementById('toast');to.textContent=msg;to.classList.add('show');setTimeout(function(){to.classList.remove('show');},2500);}
</script>
</body>
</html>
"@

    $html = $htmlBefore + $htmlAfter
    return $html
}

# ── Paths ──
 $scriptPath = $MyInvocation.PSCommandPath
if (-not $scriptPath) {
    try { $scriptPath = (Get-Item -LiteralPath $PSCommandPath -EA Stop).FullName } catch { }
}
if (-not $scriptPath) { $scriptPath = Join-Path $PWD "Generate-FolderTree.ps1" }
if (-not $OutputFile) { $OutputFile = Join-Path $PWD "FolderTree.html" }
 $OutputFile = [IO.Path]::GetFullPath($OutputFile)

# ── Status ──
if ($Status) {
    $statusDataJs = [IO.Path]::Combine([IO.Path]::GetDirectoryName($OutputFile), [IO.Path]::GetFileNameWithoutExtension($OutputFile) + ".data.js")
    if (-not (Test-Path $statusDataJs)) { Write-Host "File not found: $statusDataJs" -ForegroundColor Red; exit 1 }
    $djsRaw = Get-Content $statusDataJs -Raw -Encoding UTF8
    $m = $null
    if ($djsRaw -match 'var\s+TREE_DATA\s*=\s*') {
        $eqIdx = $djsRaw.IndexOf('=', $djsRaw.IndexOf('TREE_DATA')) + 1
        $jsonStr = $djsRaw.Substring($eqIdx).Trim()
        if ($jsonStr.EndsWith(';')) { $jsonStr = $jsonStr.Substring(0, $jsonStr.Length - 1).Trim() }
        try { $m = $jsonStr | ConvertFrom-Json } catch { $m = $null }
    }
    if (-not $m) { Write-Host "No metadata found." -ForegroundColor Yellow; exit 1 }
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "  Report Info (v2)" -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "File:              $OutputFile"
    Write-Host "Scan path:         $($m.sp)"
    Write-Host "Scan date:         $($m.sd)"
    Write-Host "Script path:       $($m.scp)"
    Write-Host "Entries:           $($m.f.PSObject.Properties.Count)"
    Write-Host "Params:            MaxDepth=$($m.p.md) Hash=$($m.p.ih) Algo=$($m.p.ha) Limit=$($m.p.hsl)MB"
    if ($m.dp -and $m.dp.Count -gt 0) {
        Write-Host "`nExcluded ($($m.dp.Count)):" -ForegroundColor Yellow
        foreach ($x in $m.dp) { Write-Host "  - $x" -ForegroundColor DarkYellow }
    } else { Write-Host "`nExcluded: none" }
    if ($m.a -and $m.a.Count -gt 0) {
        Write-Host "`nNotes ($($m.a.Count)):" -ForegroundColor Green
        foreach ($n in $m.a) {
            if ($n.p) { Write-Host "  [$($n.p)] $($n.t) ($($n.dt))" -ForegroundColor Green }
            else { Write-Host "  $($n.t) ($($n.dt))" -ForegroundColor Green }
        }
    } else { Write-Host "`nNotes: none" }
    if ($m.h -and $m.h.Count -gt 0) {
        Write-Host "`nHistory:" -ForegroundColor Magenta
        foreach ($he in $m.h) { Write-Host "  $($he.d): +$($he.a) -$($he.r) ~$($he.c)" -ForegroundColor DarkMagenta }
    }
    Write-Host ""
    exit 0
}

# ── Read existing ──
 $oldMeta = $null
 $dataJsPath = [IO.Path]::Combine([IO.Path]::GetDirectoryName($OutputFile), [IO.Path]::GetFileNameWithoutExtension($OutputFile) + ".data.js")
if (Test-Path $dataJsPath) {
    try {
        $djsRaw = Get-Content $dataJsPath -Raw -Encoding UTF8
        if ($djsRaw -match 'var\s+TREE_DATA\s*=\s*') {
            $eqIdx = $djsRaw.IndexOf('=', $djsRaw.IndexOf('TREE_DATA')) + 1
            $jsonStr = $djsRaw.Substring($eqIdx).Trim()
            if ($jsonStr.EndsWith(';')) { $jsonStr = $jsonStr.Substring(0, $jsonStr.Length - 1).Trim() }
            $oldMeta = $jsonStr | ConvertFrom-Json
        }
    } catch { $oldMeta = $null }
}
if (-not $oldMeta -and (Test-Path $OutputFile)) {
    $existingContent = Get-Content $OutputFile -Raw -Encoding UTF8
    $oldMeta = ParseMeta $existingContent
    if (-not $oldMeta) { Write-Host "Warning: no metadata in existing file, will create new." -ForegroundColor Yellow }
}

# ── Determine mode ──
 $modifyOnly = (-not $Update) -and ($DeletePattern -or $AddNote -or $RestoreAll -or $RestorePattern)
if ($modifyOnly -and -not $oldMeta) {
    Write-Host "Error: file not found or no metadata. Use -Update or create first." -ForegroundColor Red
    exit 1
}

# ── Apply modifications ──
if ($oldMeta) {
    if (-not $oldMeta.dp) { $oldMeta.dp = @() }
    if (-not $oldMeta.a)  { $oldMeta.a  = @() }
    if (-not $oldMeta.h)  { $oldMeta.h  = @() }
    if ($DeletePattern) {
        $allPaths = @($oldMeta.f.PSObject.Properties.Name)
        foreach ($pat in $DeletePattern) {
            $matched = $allPaths | Where-Object { $_ -like $pat }
            foreach ($m in $matched) {
                if ($oldMeta.dp -notcontains $m) {
                    $oldMeta.dp += $m
                    Write-Host "Excluded: $m" -ForegroundColor Yellow
                }
            }
        }
    }
    if ($RestoreAll) {
        $cnt = $oldMeta.dp.Count; $oldMeta.dp = @()
        Write-Host "Restored: $cnt entries" -ForegroundColor Green
    } elseif ($RestorePattern) {
        foreach ($pat in $RestorePattern) {
            $toRestore = @($oldMeta.dp | Where-Object { $_ -like $pat })
            foreach ($r in $toRestore) {
                $oldMeta.dp = @($oldMeta.dp | Where-Object { $_ -ne $r })
                Write-Host "Restored: $r" -ForegroundColor Green
            }
        }
    }
    if ($AddNote) {
        foreach ($note in $AddNote) {
            $nObj = @{ t=""; p=$null; dt=(Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
            if ($note -match '^(.+?)\|(.+)$') {
                $nObj.p = $matches[1].Trim(); $nObj.t = $matches[2].Trim()
                Write-Host "Note on [$($nObj.p)]: $($nObj.t)" -ForegroundColor Green
            } else {
                $nObj.t = $note.Trim()
                Write-Host "Global note: $($nObj.t)" -ForegroundColor Green
            }
            $oldMeta.a += $nObj
        }
    }
}

# ── Modify-only mode ──
if ($modifyOnly) {
    Write-Host "`nRewriting with modifications..." -ForegroundColor Cyan
    $html = BuildHTML $oldMeta @{} $null $scriptPath $OutputFile
    [IO.File]::WriteAllText($OutputFile, $html, [Text.Encoding]::UTF8)
    Write-Host "Done: $OutputFile" -ForegroundColor Green
    exit 0
}

# ── Update mode ──
if ($Update) {
    if (-not $oldMeta) {
        Write-Host "No metadata - performing full scan." -ForegroundColor Yellow
    } else {
        $scanPath = $oldMeta.sp
        $MaxDepth = $oldMeta.p.md
        $IncludeHash = [bool]$oldMeta.p.ih
        $HashAlgorithm = $oldMeta.p.ha
        $HashSizeLimitMB = $oldMeta.p.hsl
        Write-Host "=======================================" -ForegroundColor Cyan
        Write-Host "  Incremental update (v2)" -ForegroundColor Cyan
        Write-Host "=======================================" -ForegroundColor Cyan
        Write-Host "Path: $scanPath"
        if (-not (Test-Path $scanPath)) { Write-Host "Error: path not found: $scanPath" -ForegroundColor Red; exit 1 }
        Write-Host "Scanning..." -ForegroundColor DarkGray
        $scanResult = DoScan $scanPath $MaxDepth $IncludeHash $HashAlgorithm $HashSizeLimitMB
        Write-Host "Merging..." -ForegroundColor DarkGray
        $merged = MergeM $oldMeta $scanResult
        Write-Host "  New:     $($merged.Stats.N)" -ForegroundColor Green
        Write-Host "  Gone:    $($merged.Stats.G)" -ForegroundColor Red
        Write-Host "  Changed: $($merged.Stats.C)" -ForegroundColor Yellow
        Write-Host "  Excluded:$($merged.Stats.X)" -ForegroundColor DarkGray
        $html = BuildHTML $merged.Meta $merged.States $merged.Stats $scriptPath $OutputFile
        [IO.File]::WriteAllText($OutputFile, $html, [Text.Encoding]::UTF8)
        Write-Host "`nDone: $OutputFile" -ForegroundColor Green
        exit 0
    }
}

# ── Initial scan ──
if (-not $Path) { $Path = $PWD.Path }
 $scanPath = [IO.Path]::GetFullPath($Path)
if (-not (Test-Path $scanPath)) { Write-Host "Error: path not found: $scanPath" -ForegroundColor Red; exit 1 }

Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "  Scanning folder (v2)" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "Path: $scanPath"
if ($MaxDepth -gt 0) { Write-Host "Depth: $MaxDepth" }
if ($IncludeHash) { Write-Host "Hash: $HashAlgorithm (limit: $HashSizeLimitMB MB)" }
Write-Host "Scanning..." -ForegroundColor DarkGray

 $scanResult = DoScan $scanPath $MaxDepth $IncludeHash $HashAlgorithm $HashSizeLimitMB
Write-Host "Found: $($scanResult.FC) files, $($scanResult.DC) folders, $(FmtSz $scanResult.TS)" -ForegroundColor DarkGray

 $initNotes = @()
if ($AddNote) {
    foreach ($note in $AddNote) {
        $nObj = @{ t=""; p=$null; dt=(Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
        if ($note -match '^(.+?)\|(.+)$') {
            $nObj.p = $matches[1].Trim(); $nObj.t = $matches[2].Trim()
        } else { $nObj.t = $note.Trim() }
        $initNotes += $nObj
    }
}

 $filesObj = [PSCustomObject]@{}
foreach ($k in $scanResult.Files.Keys) {
    $v = $scanResult.Files[$k]
    $filesObj | Add-Member -NotePropertyName $k -NotePropertyValue ([PSCustomObject]@{ s=$v.s; lw=$v.lw; h=$v.h; d=$v.d })
}

 $meta = [PSCustomObject]@{
    sp  = $scanPath
    scp = $scriptPath
    p   = [PSCustomObject]@{ md=$MaxDepth; ih=[bool]$IncludeHash; ha=$HashAlgorithm; hsl=$HashSizeLimitMB }
    f   = $filesObj
    a   = $initNotes
    dp  = @()
    sd  = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    h   = @()
}

 $html = BuildHTML $meta @{} $null $scriptPath $OutputFile
[IO.File]::WriteAllText($OutputFile, $html, [Text.Encoding]::UTF8)

Write-Host "`nDone: $OutputFile" -ForegroundColor Green
Write-Host "To update later:" -ForegroundColor DarkGray
Write-Host "  .\Generate-FolderTree.ps1 -Update -OutputFile `"$OutputFile`"" -ForegroundColor DarkGray
