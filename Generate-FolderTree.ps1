<#
.SYNOPSIS
    Создаёт красивый HTML-документ со сворачиваемой иерархией файлов и папок.

.DESCRIPTION
    Скрипт рекурсивно обходит указанную папку и генерирует самостоятельный HTML-файл
    (без внешних зависимостей — CSS и JS встроены). Дерево можно разворачивать/сворачивать,
    искать по имени (живой поиск с подсветкой), видеть суммарный размер каждой папки,
    дату последнего изменения и доступа, а "тяжёлые" файлы/папки подсвечиваются цветом.

.PARAMETER Path
    Папка, которую нужно просканировать. По умолчанию — текущая папка.

.PARAMETER OutputFile
    Путь к результирующему HTML-файлу. По умолчанию — "FolderTree.html" рядом со скриптом.

.PARAMETER MaxDepth
    Максимальная глубина вложенности. По умолчанию без ограничений (0 = без ограничений).

.PARAMETER IncludeHash
    Если указан — вычисляет хэш каждого файла (см. -HashAlgorithm). ВНИМАНИЕ: это заметно
    замедляет работу скрипта на больших объёмах данных, так как требует чтения каждого файла
    целиком. Для очень больших файлов используйте -HashSizeLimitMB, чтобы их пропускать.

.PARAMETER HashAlgorithm
    Алгоритм хэширования: MD5, SHA1, SHA256, SHA384, SHA512. По умолчанию SHA256.
    MD5 быстрее, если хэш нужен просто для сравнения файлов, а не для криптостойкости.

.PARAMETER HashSizeLimitMB
    Файлы крупнее этого размера (в МБ) не хэшируются, чтобы не тормозить скрипт на больших
    файлах (образы дисков, видео и т.п.). По умолчанию 200 МБ. 0 = без ограничения.

.EXAMPLE
    .\Generate-FolderTree.ps1 -Path "D:\Projects" -OutputFile "C:\Temp\tree.html"

.EXAMPLE
    .\Generate-FolderTree.ps1 -Path "D:\Backup" -IncludeHash -HashAlgorithm MD5 -HashSizeLimitMB 500
#>

param(
    [Parameter(Position = 0)]
    [string]$Path = (Get-Location).Path,

    [Parameter(Position = 1)]
    [string]$OutputFile,

    [int]$MaxDepth = 0,

    [switch]$IncludeHash,

    [ValidateSet('MD5', 'SHA1', 'SHA256', 'SHA384', 'SHA512')]
    [string]$HashAlgorithm = 'SHA256',

    [int]$HashSizeLimitMB = 200
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($OutputFile)) {
    $scriptDir = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($scriptDir)) {
        if ($MyInvocation.MyCommand.Path) {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
        } else {
            $scriptDir = (Get-Location).Path
        }
    }
    $OutputFile = Join-Path $scriptDir "FolderTree.html"
}

if (-not (Test-Path -LiteralPath $Path)) {
    Write-Error "Путь не найден: $Path"
    exit 1
}

$Path = (Resolve-Path -LiteralPath $Path).Path

# Пороги "тяжести" в байтах — для файлов и для папок (суммарный размер) отдельно
$FileHeavyThresholds = [ordered]@{
    Critical = 1GB
    High     = 100MB
    Medium   = 10MB
}
$FolderHeavyThresholds = [ordered]@{
    Critical = 10GB
    High     = 1GB
    Medium   = 100MB
}

function Format-Size {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} ГБ" -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { return "{0:N2} МБ" -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { return "{0:N2} КБ" -f ($Bytes / 1KB) }
    else { return "$Bytes Б" }
}

function Get-SizeClass {
    param([long]$Bytes, [System.Collections.Specialized.OrderedDictionary]$Thresholds)
    if ($Bytes -ge $Thresholds.Critical) { return 'size-critical' }
    elseif ($Bytes -ge $Thresholds.High) { return 'size-high' }
    elseif ($Bytes -ge $Thresholds.Medium) { return 'size-medium' }
    else { return '' }
}

function Get-FileIcon {
    param([string]$Extension)
    switch ($Extension.ToLower()) {
        {$_ -in '.exe','.msi','.bat','.cmd','.ps1'} { return '⚙️' }
        {$_ -in '.zip','.rar','.7z','.tar','.gz'} { return '📦' }
        {$_ -in '.jpg','.jpeg','.png','.gif','.bmp','.svg','.webp'} { return '🖼️' }
        {$_ -in '.mp3','.wav','.flac','.ogg'} { return '🎵' }
        {$_ -in '.mp4','.avi','.mkv','.mov'} { return '🎬' }
        {$_ -in '.doc','.docx'} { return '📄' }
        {$_ -in '.xls','.xlsx','.csv'} { return '📊' }
        {$_ -in '.pdf'} { return '📕' }
        {$_ -in '.txt','.md','.log'} { return '📝' }
        {$_ -in '.iso','.vhd','.vhdx','.img'} { return '💿' }
        {$_ -in '.js','.ts','.py','.cs','.cpp','.c','.h','.java','.go','.rs','.php','.html','.css','.json','.xml','.yaml','.yml'} { return '💻' }
        default { return '📄' }
    }
}

function HtmlEncode {
    param([string]$Text)
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Build-Tree {
    param(
        [string]$FolderPath,
        [int]$Depth = 0
    )

    $sb = New-Object System.Text.StringBuilder
    $totalSize = [long]0
    $fileCountAgg = 0
    $folderCountAgg = 0

    $items = $null
    try {
        $items = Get-ChildItem -LiteralPath $FolderPath -Force -ErrorAction Stop |
            Where-Object { -not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) } |
            Sort-Object @{Expression = { $_.PSIsContainer }; Descending = $true}, Name
    } catch {
        [void]$sb.Append("<li class='error'>⚠️ Нет доступа: $(HtmlEncode $FolderPath)</li>")
        return [PSCustomObject]@{ Html = $sb.ToString(); TotalSize = [long]0; ItemCount = 0; FileCount = 0; FolderCount = 0 }
    }

    foreach ($item in $items) {
        if ($item.PSIsContainer) {
            $canExpand = ($MaxDepth -eq 0 -or $Depth -lt $MaxDepth)

            if ($canExpand) {
                $sub = Build-Tree -FolderPath $item.FullName -Depth ($Depth + 1)
                $subHtml = $sub.Html
                $subSize = $sub.TotalSize
                $childCount = $sub.ItemCount
                $fileCountAgg += $sub.FileCount
                $folderCountAgg += $sub.FolderCount + 1
            } else {
                $subHtml = "<li class='error'>… глубина ограничена</li>"
                $subSize = 0
                $childCount = 0
                try { $childCount = (Get-ChildItem -LiteralPath $item.FullName -Force -ErrorAction SilentlyContinue | Measure-Object).Count } catch {}
                $folderCountAgg += 1
            }

            $totalSize += $subSize
            $sizeClass = Get-SizeClass -Bytes $subSize -Thresholds $FolderHeavyThresholds
            $sizeLabel = Format-Size -Bytes $subSize

            [void]$sb.Append("<li class='folder $sizeClass'><details$(if($Depth -lt 1){' open'})>")
            [void]$sb.Append("<summary><span class='icon'>📁</span><span class='name'>$(HtmlEncode $item.Name)</span><span class='meta'>$childCount элем. · <strong>$sizeLabel</strong></span></summary>")
            [void]$sb.Append("<ul>$subHtml</ul>")
            [void]$sb.Append("</details></li>")
        } else {
            $icon = Get-FileIcon -Extension $item.Extension
            $size = Format-Size -Bytes $item.Length
            $modified = $item.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
            $accessed = $item.LastAccessTime.ToString("yyyy-MM-dd HH:mm")
            $sizeClass = Get-SizeClass -Bytes $item.Length -Thresholds $FileHeavyThresholds

            $hashLabel = ""
            if ($IncludeHash) {
                if ($HashSizeLimitMB -eq 0 -or $item.Length -le ($HashSizeLimitMB * 1MB)) {
                    try {
                        $hashValue = (Get-FileHash -LiteralPath $item.FullName -Algorithm $HashAlgorithm -ErrorAction Stop).Hash
                        $hashLabel = " · $HashAlgorithm`: $hashValue"
                    } catch {
                        $hashLabel = " · хэш: ошибка чтения"
                    }
                } else {
                    $hashLabel = " · хэш: файл больше $HashSizeLimitMB МБ, пропущен"
                }
            }

            $totalSize += $item.Length
            $fileCountAgg += 1

            [void]$sb.Append("<li class='file $sizeClass'><span class='icon'>$icon</span><span class='name'>$(HtmlEncode $item.Name)</span><span class='meta'>$size · изм. $modified · дост. $accessed$hashLabel</span></li>")
        }
    }

    return [PSCustomObject]@{
        Html        = $sb.ToString()
        TotalSize   = $totalSize
        ItemCount   = $items.Count
        FileCount   = $fileCountAgg
        FolderCount = $folderCountAgg
    }
}

Write-Host "Сканирую папку: $Path ..."
if ($IncludeHash) {
    Write-Host "Хэширование включено ($HashAlgorithm, лимит $HashSizeLimitMB МБ) — это может занять значительное время." -ForegroundColor Yellow
}

$treeResult = Build-Tree -FolderPath $Path
$treeHtml = $treeResult.Html
$totalFiles = $treeResult.FileCount
$totalFolders = $treeResult.FolderCount
$grandTotalSize = $treeResult.TotalSize
$grandTotalSizeLabel = Format-Size -Bytes $grandTotalSize
$generatedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

$html = @"
<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<title>Структура папки: $(HtmlEncode (Split-Path $Path -Leaf))</title>
<style>
  :root {
    --bg: #0f1117;
    --panel: #161925;
    --border: #262a3a;
    --text: #e4e6ec;
    --muted: #8b8fa3;
    --accent: #6d8dfc;
    --folder: #f5c451;
    --hover: #1d2130;
    --crit: #e06666;
    --high: #f5a623;
    --med: #f5c451;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    padding: 32px;
    background: var(--bg);
    color: var(--text);
    font-family: 'Segoe UI', Tahoma, sans-serif;
    font-size: 14px;
  }
  .header { max-width: 960px; margin: 0 auto 20px; }
  .header h1 { font-size: 20px; margin: 0 0 4px; word-break: break-all; }
  .header .path { color: var(--muted); font-size: 13px; margin-bottom: 12px; }
  .stats { display: flex; flex-wrap: wrap; gap: 16px; color: var(--muted); font-size: 13px; margin-bottom: 8px; }
  .legend { display: flex; flex-wrap: wrap; gap: 14px; color: var(--muted); font-size: 12px; margin-bottom: 12px; }
  .legend span { display: flex; align-items: center; gap: 5px; }
  .legend .dot { width: 9px; height: 9px; border-radius: 50%; display: inline-block; }
  .toolbar { max-width: 960px; margin: 0 auto 6px; display: flex; gap: 8px; align-items: center; }
  .toolbar input {
    flex: 1; padding: 8px 12px; border-radius: 8px; border: 1px solid var(--border);
    background: var(--panel); color: var(--text); font-size: 13px; outline: none;
  }
  .toolbar input:focus { border-color: var(--accent); }
  .toolbar button {
    padding: 8px 14px; border-radius: 8px; border: 1px solid var(--border);
    background: var(--panel); color: var(--text); cursor: pointer; font-size: 13px;
  }
  .toolbar button:hover { background: var(--hover); border-color: var(--accent); }
  #matchCount { color: var(--muted); font-size: 12px; white-space: nowrap; }
  .tree-container {
    max-width: 960px; margin: 10px auto 0; background: var(--panel);
    border: 1px solid var(--border); border-radius: 12px; padding: 16px 20px;
  }
  ul { list-style: none; margin: 0; padding-left: 22px; }
  .tree-container > ul { padding-left: 0; }
  li { position: relative; margin: 2px 0; }
  li.file, li.error { display: flex; align-items: center; gap: 8px; padding: 4px 8px; border-radius: 6px; border-left: 3px solid transparent; }
  li.file:hover { background: var(--hover); }
  li.error { color: #e06666; font-style: italic; }
  details { margin: 2px 0; }
  summary {
    display: flex; align-items: center; gap: 8px; padding: 4px 8px; border-radius: 6px;
    cursor: pointer; list-style: none; border-left: 3px solid transparent;
  }
  summary::-webkit-details-marker { display: none; }
  summary::before { content: '▸'; color: var(--muted); font-size: 11px; width: 10px; transition: transform 0.15s ease; }
  details[open] > summary::before { transform: rotate(90deg); }
  summary:hover { background: var(--hover); }
  .icon { flex-shrink: 0; }
  .name { flex: 1; word-break: break-all; }
  .folder > details > summary .name { color: var(--folder); font-weight: 600; }
  .meta { color: var(--muted); font-size: 12px; flex-shrink: 0; white-space: nowrap; }
  .hidden { display: none !important; }
  .highlight { background: rgba(109,141,252,0.35); border-radius: 3px; }

  .size-medium > summary, li.file.size-medium { border-left-color: var(--med); }
  .size-high > summary, li.file.size-high { border-left-color: var(--high); background: rgba(245,166,35,0.06); }
  .size-critical > summary, li.file.size-critical { border-left-color: var(--crit); background: rgba(224,102,102,0.10); }

  footer { max-width: 960px; margin: 16px auto 0; color: var(--muted); font-size: 12px; text-align: center; }
</style>
</head>
<body>
  <div class="header">
    <h1>📁 $(HtmlEncode (Split-Path $Path -Leaf))</h1>
    <div class="path">$(HtmlEncode $Path)</div>
    <div class="stats">
      <span>📄 Файлов: $totalFiles</span>
      <span>📁 Папок: $totalFolders</span>
      <span>💾 Общий размер: <strong>$grandTotalSizeLabel</strong></span>
      <span>🕒 Сгенерировано: $generatedAt</span>
    </div>
    <div class="legend">
      <span><span class="dot" style="background:var(--med)"></span>средний размер</span>
      <span><span class="dot" style="background:var(--high)"></span>крупный</span>
      <span><span class="dot" style="background:var(--crit)"></span>очень крупный</span>
    </div>
  </div>
  <div class="toolbar">
    <input type="text" id="searchBox" placeholder="Поиск по имени...">
    <span id="matchCount"></span>
    <button id="expandAll">Развернуть всё</button>
    <button id="collapseAll">Свернуть всё</button>
  </div>
  <div class="tree-container">
    <ul>
$treeHtml
    </ul>
  </div>
  <footer>Сгенерировано скриптом Generate-FolderTree.ps1</footer>

<script>
document.getElementById('expandAll').addEventListener('click', () => {
  document.querySelectorAll('details').forEach(d => d.open = true);
});
document.getElementById('collapseAll').addEventListener('click', () => {
  document.querySelectorAll('details').forEach(d => d.open = false);
});

const searchBox = document.getElementById('searchBox');
const matchCountEl = document.getElementById('matchCount');

searchBox.addEventListener('input', () => {
  const query = searchBox.value.trim().toLowerCase();
  const allLi = document.querySelectorAll('li');

  if (!query) {
    allLi.forEach(li => {
      li.classList.remove('hidden');
      const nameEl = li.querySelector(':scope > .name, :scope > details > summary > .name');
      if (nameEl) nameEl.classList.remove('highlight');
    });
    matchCountEl.textContent = '';
    return;
  }

  allLi.forEach(li => li.classList.add('hidden'));
  let matches = 0;

  allLi.forEach(li => {
    const nameEl = li.querySelector(':scope > .name, :scope > details > summary > .name');
    if (!nameEl) return;
    const text = nameEl.textContent.toLowerCase();
    if (text.includes(query)) {
      matches++;
      nameEl.classList.add('highlight');
      li.classList.remove('hidden');
      let parent = li.parentElement;
      while (parent) {
        if (parent.tagName === 'LI') {
          parent.classList.remove('hidden');
          const details = parent.querySelector(':scope > details');
          if (details) details.open = true;
        }
        if (parent.tagName === 'DETAILS') parent.open = true;
        parent = parent.parentElement;
      }
    } else {
      nameEl.classList.remove('highlight');
    }
  });

  matchCountEl.textContent = 'Найдено: ' + matches;
});
</script>
</body>
</html>
"@

$html | Out-File -FilePath $OutputFile -Encoding UTF8

Write-Host "Готово! HTML-файл создан: $OutputFile" -ForegroundColor Green
Write-Host "Файлов: $totalFiles, Папок: $totalFolders, Общий размер: $grandTotalSizeLabel"

# Автоматически открыть в браузере (уберите # в начале строки, если нужно)
# Start-Process $OutputFile
