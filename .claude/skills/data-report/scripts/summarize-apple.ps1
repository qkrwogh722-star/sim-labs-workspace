param(
  [Parameter(Mandatory)][string]$Path,
  [string]$CrosstabOut   # optional: path of the .xlsx to write for the 불량사유 x 위치 table
)
# Summarizes an apple sorting workbook (apple_A*.xlsx) without Python.
# Prints: 등급/불량사유/품종/위치 counts, numeric ranges, rule recheck, reason-total check.
# Rules mirror sim-data/templates/선별기준표.md (first match wins: 멍, 벌레, 무게<250, 당도<11).

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-SheetRows([string]$file) {
  $z = [IO.Compression.ZipFile]::OpenRead($file)
  try {
    $r = New-Object IO.StreamReader($z.GetEntry('xl/worksheets/sheet1.xml').Open())
    $x = $r.ReadToEnd(); $r.Close()
  } finally { $z.Dispose() }
  # Strings are inline (no sharedStrings.xml); strip tags from each cell.
  foreach ($m in [regex]::Matches($x, '<row [^>]*>.*?</row>')) {
    ,@([regex]::Matches($m.Value, '<c [^>]*?(?:/>|>(.*?)</c>)') |
       ForEach-Object { [regex]::Replace($_.Groups[1].Value, '<[^>]+>', '') })
  }
}

$all = @(Read-SheetRows $Path)
$h = $all[0]
$data = foreach ($i in 1..($all.Count - 1)) {
  $o = [ordered]@{}
  for ($k = 0; $k -lt $h.Count; $k++) { $o[$h[$k]] = $all[$i][$k] }
  [pscustomobject]$o
}

"파일: $Path"
"행 수: $($data.Count)   날짜: $((($data | Sort-Object 날짜 -Unique).날짜) -join ', ')"

"`n[등급]"
$data | Group-Object 등급 | ForEach-Object { "{0}: {1} ({2:P1})" -f $_.Name, $_.Count, ($_.Count / $data.Count) }

"`n[불량사유]"
$data | Group-Object 불량사유 | Sort-Object Count -Descending | ForEach-Object { "$($_.Name): $($_.Count)" }

foreach ($col in '품종', '위치') {
  "`n[$col 별 상품 비율]"
  $data | Group-Object $col | ForEach-Object {
    $g = $_.Group; $good = @($g | Where-Object 등급 -eq '상품').Count
    "{0}: n={1} 상품={2} ({3:P0})" -f $_.Name, $_.Count, $good, ($good / $_.Count)
  }
}

"`n[수치 범위]"
foreach ($c in '무게', '지름', '당도', '중심거리') {
  $m = $data | ForEach-Object { [double]$_.$c } | Measure-Object -Minimum -Maximum -Average
  "{0}: min={1} max={2} avg={3:N2}" -f $c, $m.Minimum, $m.Maximum, $m.Average
}

"`n[기준표 재판정]"
$mismatch = 0
foreach ($d in $data) {
  $re = if ($d.멍 -eq '1') { '멍' }
        elseif ($d.벌레 -eq '1') { '벌레 먹음' }
        elseif ([double]$d.무게 -lt 250) { '무게 미달' }
        elseif ([double]$d.당도 -lt 11) { '당도 미달' }
        else { '없음' }
  $grade = if ($re -eq '없음') { '상품' } else { '불량' }
  if ($re -ne $d.불량사유 -or $grade -ne $d.등급) { $mismatch++ }
}
"불일치: $mismatch"

"`n[사유별 합계 = 불량 수 확인]"
$defects = @($data | Where-Object 등급 -eq '불량').Count
$reasonSum = @($data | Where-Object 불량사유 -ne '없음').Count
"불량 등급: $defects / 사유 있는 행: $reasonSum / 일치: $($defects -eq $reasonSum)"

if ($CrosstabOut) {
  $reasons = '무게 미달', '당도 미달', '멍', '벌레 먹음', '없음'
  $locs = '가장자리', '중간', '안쪽'
  $cols = 'A', 'B', 'C', 'D', 'E'
  function S($ref, $t) { "<c r=`"$ref`" t=`"inlineStr`"><is><t>$t</t></is></c>" }
  function N($ref, $n) { "<c r=`"$ref`"><v>$n</v></c>" }
  $sb = New-Object Text.StringBuilder
  $hdr = @('불량사유') + $locs + @('합계')
  [void]$sb.Append('<row r="1">')
  for ($c = 0; $c -lt 5; $c++) { [void]$sb.Append((S "$($cols[$c])1" $hdr[$c])) }
  [void]$sb.Append('</row>')
  $rn = 2; $colTot = @(0, 0, 0, 0)
  foreach ($re in $reasons) {
    [void]$sb.Append("<row r=`"$rn`">"); [void]$sb.Append((S "A$rn" $re)); $tot = 0
    for ($c = 0; $c -lt 3; $c++) {
      $n = @($data | Where-Object { $_.불량사유 -eq $re -and $_.위치 -eq $locs[$c] }).Count
      $tot += $n; $colTot[$c] += $n; [void]$sb.Append((N "$($cols[$c+1])$rn" $n))
    }
    $colTot[3] += $tot; [void]$sb.Append((N "E$rn" $tot)); [void]$sb.Append('</row>'); $rn++
  }
  [void]$sb.Append("<row r=`"$rn`">"); [void]$sb.Append((S "A$rn" '합계'))
  for ($c = 0; $c -lt 4; $c++) { [void]$sb.Append((N "$($cols[$c+1])$rn" $colTot[$c])) }
  [void]$sb.Append('</row>')
  $ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
  $hd = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  $files = [ordered]@{
    '[Content_Types].xml' = "$hd<Types xmlns=`"http://schemas.openxmlformats.org/package/2006/content-types`"><Default Extension=`"rels`" ContentType=`"application/vnd.openxmlformats-package.relationships+xml`"/><Default Extension=`"xml`" ContentType=`"application/xml`"/><Override PartName=`"/xl/workbook.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml`"/><Override PartName=`"/xl/worksheets/sheet1.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/></Types>"
    '_rels/.rels' = "$hd<Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument`" Target=`"xl/workbook.xml`"/></Relationships>"
    'xl/workbook.xml' = "$hd<workbook $ns xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`"><sheets><sheet name=`"불량사유_위치별`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
    'xl/_rels/workbook.xml.rels' = "$hd<Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet1.xml`"/></Relationships>"
    'xl/worksheets/sheet1.xml' = "$hd<worksheet $ns><cols><col min=`"1`" max=`"1`" width=`"14`" customWidth=`"1`"/><col min=`"2`" max=`"5`" width=`"11`" customWidth=`"1`"/></cols><sheetData>$($sb.ToString())</sheetData></worksheet>"
  }
  if (Test-Path $CrosstabOut) { Remove-Item $CrosstabOut }
  New-Item -ItemType Directory -Force (Split-Path $CrosstabOut) | Out-Null
  $fs = [IO.File]::Open($CrosstabOut, 'Create')
  $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Create)
  $enc = New-Object Text.UTF8Encoding($false)
  foreach ($k in $files.Keys) {
    $s = $zip.CreateEntry($k).Open(); $b = $enc.GetBytes($files[$k]); $s.Write($b, 0, $b.Length); $s.Close()
  }
  $zip.Dispose(); $fs.Close()
  "`n교차표 저장: $CrosstabOut"
}
