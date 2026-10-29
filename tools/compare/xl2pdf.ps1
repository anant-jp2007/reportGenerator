param([string]$in, [string]$out)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($in, 0, $true)
  $ws = $wb.Worksheets.Item(1)
  "zoom=" + $ws.PageSetup.Zoom + " fitW=" + $ws.PageSetup.FitToPagesWide
  $cols = @(); for ($c = 1; $c -le 13; $c++) { $cols += [string]$ws.Columns.Item($c).Width }
  "colpts=" + ($cols -join ",")
  $rows = @(); for ($r = 1; $r -le 42; $r++) { $rows += [string]$ws.Rows.Item($r).Height }
  "rowpts=" + ($rows -join ",")
  $wb.ExportAsFixedFormat(0, $out)
  $wb.Close($false)
} finally { $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
