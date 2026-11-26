param(
    [Parameter(Mandatory = $true)][string]$Path,
    [int]$TopN = 10
)
# Read-only profile of an .xlsx file: sheets, header, row count, and per-column stats.
# Opens the workbook with ZipFile.OpenRead only; never writes anything.
# Handles inline strings (<is><t>) and shared strings (xl/sharedStrings.xml).
# Keep this file ASCII-only: Windows PowerShell 5.1 misreads BOM-less UTF-8 literals.

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$inv = [Globalization.CultureInfo]::InvariantCulture

function Read-Entry($zip, [string]$name) {
    $e = $zip.GetEntry($name)
    if (-not $e) { return $null }
    $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
    try { return $sr.ReadToEnd() } finally { $sr.Close() }
}

function Get-ColIndex([string]$ref) {
    $letters = ($ref -replace '[0-9]', '').ToUpper()
    $n = 0
    foreach ($ch in $letters.ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }
    return $n - 1
}

function Format-Num([double]$d) {
    if ([double]::IsNaN($d)) { return 'NaN' }
    return [math]::Round($d, 4).ToString($inv)
}

$full = (Resolve-Path -LiteralPath $Path).Path
$zip = [IO.Compression.ZipFile]::OpenRead($full)
try {
    # Shared strings (optional).
    $shared = @()
    $ssXml = Read-Entry $zip 'xl/sharedStrings.xml'
    if ($ssXml) {
        [xml]$ss = $ssXml
        $shared = @($ss.sst.si | ForEach-Object { ($_.SelectNodes('.//*[local-name()="t"]') | ForEach-Object { $_.InnerText }) -join '' })
    }

    # Sheet names -> part paths via workbook.xml + rels.
    [xml]$wb = Read-Entry $zip 'xl/workbook.xml'
    [xml]$rels = Read-Entry $zip 'xl/_rels/workbook.xml.rels'
    $relMap = @{}
    foreach ($r in $rels.Relationships.Relationship) { $relMap[$r.Id] = $r.Target }
    $sheets = @()
    foreach ($s in $wb.workbook.sheets.sheet) {
        $rid = $s.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $target = $relMap[$rid] -replace '^/?xl/', '' -replace '^/', ''
        $sheets += [pscustomobject]@{ Name = $s.name; Part = 'xl/' + $target }
    }

    "FILE: $full"
    "SIZE_BYTES: $((Get-Item -LiteralPath $full).Length)"
    "SHARED_STRINGS: $(if ($ssXml) { 'yes (' + $shared.Count + ')' } else { 'no (inline strings)' })"
    "SHEETS: $(($sheets | ForEach-Object { $_.Name }) -join ', ')"

    foreach ($sh in $sheets) {
        ''
        "=== SHEET: $($sh.Name) ($($sh.Part)) ==="
        [xml]$x = Read-Entry $zip $sh.Part
        $rowNodes = @($x.worksheet.sheetData.row)
        if ($rowNodes.Count -eq 0) { 'EMPTY'; continue }

        # Build a grid of cell text keyed by column index.
        $grid = New-Object System.Collections.Generic.List[object]
        $maxCol = -1
        foreach ($row in $rowNodes) {
            $cells = @{}
            foreach ($c in @($row.c)) {
                if (-not $c) { continue }
                $ci = Get-ColIndex $c.r
                if ($ci -gt $maxCol) { $maxCol = $ci }
                $t = $c.GetAttribute('t')
                $val = $null
                if ($t -eq 'inlineStr') { $val = ($c.SelectNodes('.//*[local-name()="t"]') | ForEach-Object { $_.InnerText }) -join '' }
                elseif ($t -eq 's') { $val = $shared[[int]$c.v] }
                else { $val = [string]$c.v }
                $cells[$ci] = $val
            }
            $grid.Add($cells)
        }

        $header = $grid[0]
        $data = $grid | Select-Object -Skip 1
        $nRows = @($data).Count
        "ROWS (excluding header): $nRows"
        "COLUMNS: $($maxCol + 1)"
        ''
        for ($ci = 0; $ci -le $maxCol; $ci++) {
            $name = [string]$header[$ci]
            if (-not $name) { $name = "(col $ci, no header)" }
            $vals = @($data | ForEach-Object { $_[$ci] })
            $nonEmpty = @($vals | Where-Object { $_ -ne $null -and [string]$_ -ne '' })
            $empty = $nRows - $nonEmpty.Count
            $nums = New-Object System.Collections.Generic.List[double]
            $allNum = $nonEmpty.Count -gt 0
            foreach ($v in $nonEmpty) {
                $d = 0.0
                if ([double]::TryParse([string]$v, [Globalization.NumberStyles]::Float, $inv, [ref]$d)) { $nums.Add($d) } else { $allNum = $false; break }
            }
            $distinct = @($nonEmpty | Group-Object | Sort-Object -Property @{Expression='Count';Descending=$true}, @{Expression='Name';Descending=$false})
            if ($allNum) {
                $sorted = $nums | Sort-Object
                $mid = [int][math]::Floor(($sorted.Count - 1) / 2)
                $median = if ($sorted.Count % 2 -eq 1) { $sorted[$mid] } else { ($sorted[$mid] + $sorted[$mid + 1]) / 2 }
                $mean = ($nums | Measure-Object -Average).Average
                $sd = [math]::Sqrt((($nums | ForEach-Object { ($_ - $mean) * ($_ - $mean) }) | Measure-Object -Sum).Sum / [math]::Max(1, $nums.Count - 1))
                "- [$ci] $name | NUMERIC | n=$($nonEmpty.Count) empty=$empty distinct=$($distinct.Count) | min=$(Format-Num $sorted[0]) max=$(Format-Num $sorted[-1]) mean=$(Format-Num $mean) median=$(Format-Num $median) sd=$(Format-Num $sd)"
                if ($distinct.Count -le $TopN) {
                    "    values: " + (($distinct | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', ')
                }
            } else {
                "- [$ci] $name | TEXT | n=$($nonEmpty.Count) empty=$empty distinct=$($distinct.Count)"
                "    top: " + (($distinct | Select-Object -First $TopN | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', ') + $(if ($distinct.Count -gt $TopN) { ", ... (+$($distinct.Count - $TopN) more)" } else { '' })
            }
        }
    }
}
finally { $zip.Dispose() }
