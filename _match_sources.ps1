# Match sub-source rows (PQ-58 12, PQ-58 13, VM-88) against master HK + Changed parcels.
$ErrorActionPreference='Stop'

$f = 'C:\Users\Rustam\Desktop\ПҚ-58 ва ВМ-88 ф.xlsm'
$sd  = Get-Content -LiteralPath 'C:\Users\Rustam\Desktop\Ер бўйича сайт\sections_data.js' -Raw -Encoding utf8
$sdJson = $sd -replace '^\s*window\.SECTIONS_RAW_FULL\s*=\s*', '' -replace ';\s*$', ''
$master = $sdJson | ConvertFrom-Json

# Coord-normalisation char classes
$RX_PRIME    = "[$([char]0x02B9)$([char]0x02BC)$([char]0x2032)$([char]0x00B4)$([char]0x2019)$([char]0x02BB)$([char]0x2018)]"
$RX_DPRIME   = "[$([char]0x02BA)$([char]0x2033)$([char]0x201D)$([char]0x201C)]"
$RX_DEGREE   = "[$([char]0x2070)$([char]0x00BA)$([char]0x00B0)]"

function Norm-CoordText { param([string]$txt)
  if(-not $txt){ return '' }
  $t = $txt
  $t = $t -replace $RX_DEGREE, [char]0x00B0
  $t = $t -replace $RX_PRIME, "'"
  $t = $t -replace $RX_DPRIME, '"'
  $t = $t -replace "`t", ' '
  $t = $t -replace '[Сс]', 'N'
  $t = $t -replace '[Юю]', 'S'
  $t = $t -replace '[Вв]', 'E'
  $t = $t -replace '[Зз]', 'W'
  $t = $t -replace '(\d),(\d)', '$1.$2'
  return $t
}

function To-Decimal { param([double]$deg, [double]$min, [double]$sec, [string]$dir)
  $v = $deg + $min/60 + $sec/3600
  if($dir -eq 'S' -or $dir -eq 'W'){ $v = -$v }
  return $v
}

function Extract-FirstLatLon { param([string]$txt)
  if(-not $txt){ return $null }
  $t = Norm-CoordText $txt
  $rx = '(\d{1,3})\s*' + [char]0x00B0 + '\s*(\d{1,2})?\s*' + "'" + '?\s*(\d{1,2}(?:\.\d+)?)?\s*"?\s*([NSEW])\s*[, ]?\s*(\d{1,3})\s*' + [char]0x00B0 + '\s*(\d{1,2})?\s*' + "'" + '?\s*(\d{1,2}(?:\.\d+)?)?\s*"?\s*([NSEW])'
  $m = [regex]::Match($t, $rx)
  if($m.Success){
    $a_deg=[double]$m.Groups[1].Value; $a_min=0.0; $a_sec=0.0; $a_dir=$m.Groups[4].Value
    if($m.Groups[2].Value){ $a_min = [double]$m.Groups[2].Value }
    if($m.Groups[3].Value){ $a_sec = [double]$m.Groups[3].Value }
    $b_deg=[double]$m.Groups[5].Value; $b_min=0.0; $b_sec=0.0; $b_dir=$m.Groups[8].Value
    if($m.Groups[6].Value){ $b_min = [double]$m.Groups[6].Value }
    if($m.Groups[7].Value){ $b_sec = [double]$m.Groups[7].Value }
    if($a_dir -eq 'N' -or $a_dir -eq 'S'){
      return [pscustomobject]@{ lat=(To-Decimal $a_deg $a_min $a_sec $a_dir); lon=(To-Decimal $b_deg $b_min $b_sec $b_dir) }
    } else {
      return [pscustomobject]@{ lat=(To-Decimal $b_deg $b_min $b_sec $b_dir); lon=(To-Decimal $a_deg $a_min $a_sec $a_dir) }
    }
  }
  return $null
}

function Norm-Contour { param([string]$s)
  if(-not $s){ return '' }
  $t = $s.ToLower()
  $t = $t -replace [char]0x049B, 'q'
  $t = $t -replace 'қ', 'q'
  $t = $t -replace '\s+', ''
  return $t.Trim()
}

function Get-Field { param($row, [string[]]$keys)
  foreach($k in $keys){
    if($row.PSObject.Properties.Match($k).Count){
      $v = "$($row.$k)"
      if($v.Trim()){ return $v }
    }
  }
  # Substring fallback
  foreach($key in $keys){
    foreach($p in $row.PSObject.Properties){
      if($p.Name.ToLower().Contains($key.ToLower())){
        $v = "$($p.Value)"
        if($v.Trim()){ return $v }
      }
    }
  }
  return ''
}

function Get-CoordValue { param($row)
  foreach($p in $row.PSObject.Properties){
    if($p.Name.ToLower().Contains('координ')){
      $v = "$($p.Value)"
      if($v.Trim()){ return $v }
    }
  }
  return ''
}

function Get-ContourValue { param($row)
  foreach($p in $row.PSObject.Properties){
    if($p.Name.ToLower().Contains('контур')){
      $v = "$($p.Value)"
      if($v.Trim()){ return $v }
    }
  }
  return ''
}

# Build master indices
$contourIdx = @{}
$coordIdx = New-Object 'System.Collections.Generic.List[psobject]'
foreach($sec in @('hk','changed')){
  foreach($r in $master.sections.$sec){
    $rid = $r._row
    $contour = Get-ContourValue $r
    if($contour){
      $parts = (Norm-Contour $contour) -split '[,;|/]'
      foreach($c in $parts){
        if($c){
          if(-not $contourIdx.ContainsKey($c)){ $contourIdx[$c] = New-Object 'System.Collections.Generic.List[psobject]' }
          [void]$contourIdx[$c].Add([pscustomobject]@{ section=$sec; rowId=$rid })
        }
      }
    }
    $coordTxt = Get-CoordValue $r
    if($coordTxt){
      $ll = Extract-FirstLatLon $coordTxt
      if($ll){
        [void]$coordIdx.Add([pscustomobject]@{ lat=$ll.lat; lon=$ll.lon; section=$sec; rowId=$rid })
      }
    }
  }
}
Write-Output ("Master contours indexed: {0} keys" -f $contourIdx.Count)
Write-Output ("Master coords indexed:   {0} entries" -f $coordIdx.Count)

# Read sub-files DIRECTLY from xlsm so we can implement the district-inheritance pattern
$excel = New-Object -ComObject Excel.Application
$excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.AutomationSecurity=3
$wb = $excel.Workbooks.Open($f, 0, $true)

function Read-SheetWithInheritedDistrict {
  param($wsName, [int]$dataStart, [int[]]$headerRows)
  $ws = $wb.Worksheets.Item($wsName)
  $cols = $ws.UsedRange.Columns.Count
  $lastRow = $ws.UsedRange.Rows.Count
  # Build headers
  $headerSets=@()
  foreach($hr in $headerRows){ $headerSets += ,($ws.Range($ws.Cells($hr,1), $ws.Cells($hr,$cols)).Value2) }
  $headers = New-Object 'System.Collections.Generic.List[string]'
  for($i=1;$i -le $cols;$i++){
    $parts=@()
    foreach($set in $headerSets){
      $v = ("$($set[1,$i])" -replace "`r"," " -replace "`n"," ").Trim()
      if($v){ $parts += $v }
    }
    $hdr = if($parts.Count){ ($parts -join ' / ') } else { "Col $i" }
    [void]$headers.Add($hdr)
  }
  $rng = $ws.Range($ws.Cells($dataStart,1), $ws.Cells($lastRow,$cols)).Value2
  $rows = New-Object 'System.Collections.Generic.List[psobject]'
  $currentDistrict = ''
  $realRows = $lastRow - ($dataStart - 1)
  for($r=1;$r -le $realRows;$r++){
    $no = "$($rng[$r,1])".Trim()
    $col2 = "$($rng[$r,2])".Trim()
    # District-group separator: row where col1 is a Cyrillic district label (e.g. "Кармана тумани") and col 2+ are mostly empty
    if(($no -match 'тумани|шаҳри|шаҳар') -and (-not $col2 -or $col2.Length -lt 2)){
      $currentDistrict = $no
      continue
    }
    # Skip totals rows
    if($no -match '^Жами' -or $col2 -match '^Жами' -or $no -match '^ЖАМИ'){ continue }
    # Need at least a row number OR a district hint
    if(-not $no -and -not $col2){ continue }
    $obj = [ordered]@{}
    $obj['_row'] = ($r + $dataStart - 1)
    $obj['_districtInherited'] = $currentDistrict
    for($i=1;$i -le $cols;$i++){
      $v = $rng[$r,$i]
      if($null -eq $v){ continue }
      if($v -is [string]){
        $s = ($v -replace "`r","`n").Trim()
        if($s.Length -eq 0){ continue }
        $obj[$headers[$i-1]] = $s
      } else {
        $obj[$headers[$i-1]] = $v
      }
    }
    [void]$rows.Add([pscustomobject]$obj)
  }
  return $rows
}

$pq12 = Read-SheetWithInheritedDistrict '02_Номма-ном' 7 @(5,6)
$pq13 = Read-SheetWithInheritedDistrict '06_Номма-ном (2)' 6 @(4)
$vm88 = Read-SheetWithInheritedDistrict '09_Номма-ном' 6 @(4)

$wb.Close($false); $excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[gc]::Collect(); [gc]::WaitForPendingFinalizers()

Write-Output ("PQ-12 rows: {0}" -f $pq12.Count)
Write-Output ("PQ-13 rows: {0}" -f $pq13.Count)
Write-Output ("VM-88 rows: {0}" -f $vm88.Count)

function Match-Row { param($row, [bool]$allowSplit)
  $matches = New-Object 'System.Collections.Generic.List[psobject]'
  $contour = Get-ContourValue $row
  if($contour){
    foreach($c in (Norm-Contour $contour) -split '[,;|/]'){
      if($c -and $contourIdx.ContainsKey($c)){
        foreach($hit in $contourIdx[$c]){ [void]$matches.Add($hit) }
      }
    }
  }
  if($matches.Count -gt 0 -and -not $allowSplit){ return $matches }
  $coordTxt = Get-CoordValue $row
  if($coordTxt){
    $ll = Extract-FirstLatLon $coordTxt
    if($ll){
      foreach($entry in $coordIdx){
        $dl = [Math]::Abs($entry.lat - $ll.lat)
        $do = [Math]::Abs($entry.lon - $ll.lon)
        if($dl -lt 0.002 -and $do -lt 0.002){
          [void]$matches.Add([pscustomobject]@{ section=$entry.section; rowId=$entry.rowId })
        }
      }
    }
  }
  return $matches
}

$tagMap = @{}
$stats = @{
  pq12 = @{ rows=0; matched=0; multi=0; unmatched=0 }
  pq13 = @{ rows=0; matched=0; multi=0; unmatched=0 }
  vm88 = @{ rows=0; matched=0; multi=0; unmatched=0 }
}
$unmatched = @{
  pq12 = New-Object 'System.Collections.Generic.List[psobject]'
  pq13 = New-Object 'System.Collections.Generic.List[psobject]'
  vm88 = New-Object 'System.Collections.Generic.List[psobject]'
}

$datasets = @{ pq12=$pq12; pq13=$pq13; vm88=$vm88 }
foreach($source in @('pq12','pq13','vm88')){
  foreach($r in $datasets[$source]){
    $stats[$source].rows++
    $allowSplit = ($source -eq 'pq13')
    $hits = Match-Row $r $allowSplit
    if($hits.Count -eq 0){
      $stats[$source].unmatched++
      $coordVal = (Get-CoordValue $r)
      $coordVal = $coordVal -replace '\s+', ' '
      $cs = if($coordVal.Length -gt 60){ $coordVal.Substring(0,60) } else { $coordVal }
      $sample = [pscustomobject]@{
        district = "$($r._districtInherited)"
        row = $r._row
        contour = (Get-ContourValue $r)
        ha = (Get-Field $r @('Гектар','Жами майдон (га)','Жами ер майдон (гектар)'))
        coord = $cs
      }
      [void]$unmatched[$source].Add($sample)
      continue
    }
    $stats[$source].matched++
    if($hits.Count -gt 1){ $stats[$source].multi++ }
    foreach($h in $hits){
      $key = "$($h.section)-$($h.rowId)"
      if(-not $tagMap.ContainsKey($key)){ $tagMap[$key] = New-Object 'System.Collections.Generic.List[string]' }
      if(-not $tagMap[$key].Contains($source)){ [void]$tagMap[$key].Add($source) }
    }
  }
}

Write-Output ""
Write-Output "=== Matching audit ==="
foreach($s in @('pq12','pq13','vm88')){
  $st = $stats[$s]
  Write-Output ("{0}: rows={1} matched={2} multi={3} unmatched={4}" -f $s, $st.rows, $st.matched, $st.multi, $st.unmatched)
}
Write-Output ("Total tagged master parcels: {0}" -f $tagMap.Count)
Write-Output ""
foreach($s in @('pq12','pq13','vm88')){
  if($unmatched[$s].Count -gt 0){
    Write-Output ("--- Unmatched {0} (first 5) ---" -f $s)
    foreach($u in ($unmatched[$s] | Select-Object -First 5)){
      Write-Output ("  R{0} district='{1}' contour='{2}' ha='{3}' coord='{4}'" -f $u.row, $u.district, $u.contour, $u.ha, $u.coord)
    }
  }
}

$idx = @{}
foreach($k in $tagMap.Keys){ $idx[$k] = $tagMap[$k] }
$idxJson = $idx | ConvertTo-Json -Compress -Depth 4
[IO.File]::WriteAllText('C:\Users\Rustam\Desktop\Ер бўйича сайт\source_tags_index.json',
  $idxJson, [Text.UTF8Encoding]::new($false))
Write-Output ""
Write-Output ("Wrote source_tags_index.json with {0} entries" -f $tagMap.Count)
