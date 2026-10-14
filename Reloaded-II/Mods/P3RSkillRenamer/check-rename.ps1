$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Xml.Linq

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$rep = New-Object System.Collections.Generic.List[string]

function Say([string]$m, [string]$color) {
    $rep.Add($m)
    if ($color) { Write-Host $m -ForegroundColor $color }
    else { Write-Host $m }
}

$Sheet = Join-Path $here '技能改名.xlsx'
if (-not (Test-Path $Sheet)) {
    Say "没找到 $Sheet" 'Red'
    Say "请把 技能改名.xlsx 放到本文件所在目录后 ，再运行检查。" 'Red'
    exit 2
}

Say "读取: $Sheet" 'Cyan'
Say ""

# ---------------- parse xlsx (zip of xml, no external libs) ----------------
$tmp = Join-Path $env:TEMP ("p3rchk_" + [Guid]::NewGuid().ToString('N') + ".xlsx")
try {
    $fsrc = New-Object System.IO.FileStream($Sheet, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
    $fdst = New-Object System.IO.FileStream($tmp, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, ([System.IO.FileShare]::None))
    $fsrc.CopyTo($fdst); $fdst.Close(); $fsrc.Close()

    $zip = [System.IO.Compression.ZipFile]::OpenRead($tmp)
    $ns  = [System.Xml.Linq.XNamespace]::Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main")

    $shared = @()
    $e = $zip.GetEntry("xl/sharedStrings.xml")
    if ($e) {
        $rd = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
        $xd = [System.Xml.Linq.XDocument]::Parse($rd.ReadToEnd()); $rd.Close()
        foreach ($si in $xd.Root.Descendants($ns + "si")) {
            $b = New-Object System.Text.StringBuilder
            foreach ($t in $si.Descendants($ns + "t")) { [void]$b.Append($t.Value) }
            $shared += $b.ToString()
        }
    }

    $se = $zip.Entries | Where-Object { $_.FullName -like "xl/worksheets/*.xml" } | Select-Object -First 1
    if (-not $se) { throw "xlsx 里没有工作表" }
    $rd2 = New-Object System.IO.StreamReader($se.Open(), [System.Text.Encoding]::UTF8)
    $doc = [System.Xml.Linq.XDocument]::Parse($rd2.ReadToEnd()); $rd2.Close()
    $zip.Dispose()

    $rows = @()
    foreach ($row in $doc.Descendants($ns + "row")) {
        $cells = @{}
        foreach ($c in $row.Elements($ns + "c")) {
            $a = $c.Attribute("r"); if (-not $a) { continue }
            $col = 0
            foreach ($ch in $a.Value.ToCharArray()) {
                if ($ch -ge [char]65 -and $ch -le [char]90) { $col = $col * 26 + ([int][char]$ch - 64) } else { break }
            }
            $col = $col - 1
            if ($col -lt 0 -or $col -gt 2) { continue }
            $ta = $c.Attribute("t"); $t = if ($ta) { $ta.Value } else { "n" }
            if ($t -eq 's') {
                $ve = $c.Element($ns + "v")
                $ix = [int]$ve.Value
                $val = if ($ix -lt $shared.Count) { $shared[$ix] } else { "" }
            } elseif ($t -eq 'inlineStr') {
                $is = $c.Element($ns + "is")
                $val = if ($is) { ($is.Descendants($ns + "t") | ForEach-Object { $_.Value }) -join '' } else { "" }
            } else {
                $ve = $c.Element($ns + "v"); $val = if ($ve) { $ve.Value } else { "" }
            }
            $cells[$col] = [string]$val
        }
        $rows += , @($cells[0], $cells[1], $cells[2])
    }
}
finally { try { Remove-Item $tmp -Force -ErrorAction SilentlyContinue } catch {} }

if ($rows.Count -eq 0) {
    Say "表里没有数据。" 'Red'
    exit 3
}

# ---------------- header ----------------
$headerMap = @{ 'id'=$true; '原名'=$true; '中文名'=$true; '名称'=$true; '技能名'=$true; '旧名'=$true; 'name'=$true; 'skill'=$true
               '改为'=$true; '新名'=$true; '新名称'=$true; '改成'=$true; 'rename'=$true; 'to'=$true; 'new'=$true }
$mode = 'ABC'; $idxA = 0; $idxOld = 1; $idxNew = 2
$h0 = ([string]$rows[0][0]).Trim().ToLower(); $h2 = ([string]$rows[0][2]).Trim().ToLower()
if ($headerMap.ContainsKey($h0) -and $headerMap.ContainsKey($h2)) {
    Say ("表头: A=$($rows[0][0])  B=$($rows[0][1])  C=$($rows[0][2])   (ID + 原名 + 改为)") 'Green'
} else {
    $mode = 'AB'; $idxA = -1; $idxOld = 0; $idxNew = 1
    Say ("表头: A=$($rows[0][0])  B=$($rows[0][1])   (原名 + 改为，按名字自动查找索引)") 'Green'
}

# ---------------- reference table ----------------
$ref = @{}
$refFile = Join-Path $here 'SkillNameTable.tsv'
if (Test-Path $refFile) {
    foreach ($line in [System.IO.File]::ReadAllLines($refFile, [System.Text.Encoding]::UTF8)) {
        if ($line -match '^[#\s]' -or $line.Length -eq 0) { continue }
        if ($line -match '^index\tvalue') { continue }
        $i = $line.IndexOf("`t"); if ($i -lt 0) { continue }
        $k = 0; if (-not [int]::TryParse($line.Substring(0,$i), [ref]$k)) { continue }
        $ref[$k] = $line.Substring($i+1)
    }
    Say ("参考表: {0} 条 (SkillNameTable.tsv)" -f $ref.Count)
} else {
    Say "参考表: 没有找到 SkillNameTable.tsv，跳过原名校验" 'Yellow'
}

# ---------------- validate ----------------
$err = 0; $warn = 0; $ok = 0
$seenId = @{}

for ($r = 1; $r -lt $rows.Count; $r++) {
    $row = $rows[$r]
    $lineNo = $r + 1
    $old = ([string]$row[$idxOld]).Trim()
    $new = ([string]$row[$idxNew]).Trim()
    $idS = if ($idxA -ge 0) { ([string]$row[$idxA]).Trim() } else { "" }

    if ($new.Length -eq 0) { continue }
    if ($old.Length -eq 0 -and $mode -eq 'AB') {
        Say ("  第{0}行: 新名 '{1}' 没有对应的原名，跳过" -f $lineNo, $new) 'Yellow'; $warn++; continue
    }

    $id = -1
    if ($mode -eq 'ABC') {
        if ($idS -match '^\d+$') {
            $id = [int]$idS
            if ($ref.Count -gt 0 -and $id -ge $ref.Count) {
                Say ("  第{0}行: ID {1} 超出范围 (0..{2})" -f $lineNo, $id, ($ref.Count-1)) 'Red'; $err++; continue
            }
            if ($seenId.ContainsKey($id)) {
                Say ("  第{0}行: ID {1} 重复 (第{2}行已用)" -f $lineNo, $id, $seenId[$id]) 'Red'; $err++; continue
            }
            $seenId[$id] = $lineNo
            if ($ref.Count -gt 0 -and $old.Length -gt 0) {
                if ((([string]$ref[$id]) -replace '\s','') -ne ($old -replace '\s','')) {
                    Say ("  第{0}行: ID {1} 处的实际原名是 '{2}'，表里写的是 '{3}'" -f $lineNo, $id, $ref[$id], $old) 'Yellow'; $warn++
                }
            }
        }
        else {
            Say ("  第{0}行: ID 栏 '{1}' 不是数字，将改为按原名查找" -f $lineNo, $idS) 'Yellow'; $warn++
        }
    }

    $ok++
}

Say ""
Say ("有效改名: {0} 条    警告: {1}    错误: {2}" -f $ok, $warn, $err) $(if ($err -gt 0) { 'Red' } else { 'Green' })
if ($ok -eq 0) { Say "没有可生效的改名 (C 列是不是都空着？)" 'Yellow' }

Say ""
Say "自查完成。" 'DarkGray'

if ($err -gt 0) { exit 1 }
exit 0