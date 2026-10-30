<#
.SYNOPSIS
    Golden-master test: runs the update in InstallationTracker.xlsm and compares every
    status, issue count and dashboard figure with data\expected.json.

.DESCRIPTION
    expected.json is written by tools\generate_demo_data.py, which calculates the figures
    in Python, separately from the VBA code. The test also checks that notes typed in the
    issue sheets survive an update and that notes of closed issues go to Notes Archive.
    The workbook is closed without saving.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\test_tracker.ps1
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$expected = Get-Content (Join-Path $root 'data\expected.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$script:passed = 0
$script:failed = 0

function Check([string]$name, $actual, $want) {
    if ([double]$actual -eq [double]$want) {
        $script:passed++
    } else {
        $script:failed++
        Write-Host ("FAIL  {0}: workbook {1}, expected {2}" -f $name, $actual, $want)
    }
}

# Counts the values of one table column
function CountColumn($table, [string]$column) {
    $counts = @{}
    if ($table.ListRows.Count -eq 0) { return $counts }
    $values = $table.ListColumns.Item($column).DataBodyRange.Value2
    if ($table.ListRows.Count -eq 1) { $counts[[string]$values] = 1; return $counts }
    foreach ($v in $values) { $counts[[string]$v] = 1 + [int]$counts[[string]$v] }
    return $counts
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open((Join-Path $root 'InstallationTracker.xlsm'))
    $excel.Run('RunUpdate', $false)
    $excel.CalculateFull()
    $dash = $wb.Worksheets.Item('Dashboard')
    Write-Host ($dash.Range('B2').Text)

    # 1. Order statuses
    $orders = $wb.Worksheets.Item('Orders').ListObjects.Item('tblOrders')
    Check 'orders' $orders.ListRows.Count $expected.orders
    $status = CountColumn $orders 'Status'
    foreach ($p in $expected.status.PSObject.Properties) { Check "status '$($p.Name)'" ([int]$status[$p.Name]) $p.Value }

    # 2. Issues on the three issue sheets
    $issues = @{}
    foreach ($t in @(@('Schedule Issues', 'tblScheduleIssues'), @('Spec Issues', 'tblSpecIssues'), @('Contract Issues', 'tblContractIssues'))) {
        $c = CountColumn $wb.Worksheets.Item($t[0]).ListObjects.Item($t[1]) 'Issue'
        foreach ($k in $c.Keys) { $issues[$k] = $c[$k] + [int]$issues[$k] }
    }
    foreach ($p in $expected.issues.PSObject.Properties) { Check "issue '$($p.Name)'" ([int]$issues[$p.Name]) $p.Value }

    # 3. Monthly table C14:J25 (column F, % of plan, is derived and not compared)
    $grid = $dash.Range('C14:J25').Value2
    $columns = @{ plan = 1; sold = 2; installed = 3; specified = 5; delayed = 6; open = 7; no_manager = 8 }
    for ($m = 0; $m -lt 12; $m++) {
        $row = $expected.monthly[$m]
        foreach ($k in $columns.Keys) { Check "$($row.month) $k" $grid[($m + 1), $columns[$k]] $row.$k }
    }

    # 4. Products, year to date (B36:G40)
    $prod = $dash.Range('B36:G40').Value2
    for ($r = 1; $r -le 5; $r++) {
        $want = $expected.products.($prod[$r, 1])
        if ($null -eq $want) { $script:failed++; Write-Host "FAIL  unexpected product '$($prod[$r, 1])'"; continue }
        Check "$($prod[$r, 1]) plan" $prod[$r, 2] $want.plan
        Check "$($prod[$r, 1]) installed" $prod[$r, 3] $want.installed
        Check "$($prod[$r, 1]) open" $prod[$r, 5] $want.open
        Check "$($prod[$r, 1]) delayed" $prod[$r, 6] $want.delayed
    }

    # 5. KPI tiles
    Check 'KPI plan YTD' $dash.Range('B10').Value2 $expected.kpi.plan_ytd
    Check 'KPI installed YTD' $dash.Range('D10').Value2 $expected.kpi.installed_ytd
    Check 'KPI open orders YTD' $dash.Range('H10').Value2 $expected.kpi.open_ytd
    Check 'KPI delayed start' $dash.Range('J10').Value2 $expected.kpi.delayed
    Check 'KPI open issues' $dash.Range('L10').Value2 $expected.kpi.open_issues

    # 6. Notes: one typed on a real issue must stay, one on a closed issue must be archived
    $sched = $wb.Worksheets.Item('Schedule Issues').ListObjects.Item('tblScheduleIssues')
    $noteCol = $sched.ListColumns.Item('Note').Index
    $keyOfFirst = $sched.DataBodyRange.Cells.Item(1, 1).Value2
    $sched.DataBodyRange.Cells.Item(1, $noteCol).Value2 = 'Called the crew leader'
    $fake = $sched.ListRows.Add()
    $fake.Range.Cells.Item(1, 1).Value2 = 'TEST|closed issue'
    $fake.Range.Cells.Item(1, $noteCol).Value2 = 'Note on an issue that is gone'
    $excel.Run('RunUpdate', $false)

    $kept = $false
    $keys = $sched.ListColumns.Item('Key').DataBodyRange.Value2
    $notes = $sched.ListColumns.Item('Note').DataBodyRange.Value2
    for ($r = 1; $r -le $sched.ListRows.Count; $r++) {
        if ($keys[$r, 1] -eq $keyOfFirst -and $notes[$r, 1] -eq 'Called the crew leader') { $kept = $true }
    }
    Check 'note kept after update' ([int]$kept) 1
    $archive = CountColumn $wb.Worksheets.Item('Notes Archive').ListObjects.Item('tblNotesArchive') 'Key'
    Check 'closed note archived' ([int]$archive['TEST|closed issue']) 1
    Check 'no fake row left' ([int]((CountColumn $sched 'Key')['TEST|closed issue'])) 0

    $wb.Close($false)
}
finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ("{0} checks passed, {1} failed" -f $script:passed, $script:failed)
if ($script:failed -gt 0) { exit 1 }
