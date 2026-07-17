#Requires -Version 7
<#
.SYNOPSIS
    Microsoft-only (Excel COM) replacements for the ImportExcel/EPPlus operations used by the WARA
    report generator (3_wara_reports_generator.ps1). Removes the third-party ImportExcel dependency
    from Phase 3 so the report can run in disconnected / sovereign environments with only desktop
    Microsoft Excel installed.

.DESCRIPTION
    Provides:
      - Import-WARAExcelViaCom / Initialize-WARAExpertAnalysisCache : read Expert-Analysis worksheets
        into PSCustomObjects (drop-in for Import-Excel -WorksheetName -StartRow).
      - Build-WARAAssessmentFindingsCom : copy the Assessment-Findings template, write the data sheets
        (3.ImpactedResources / 2.Recommendations / 6.WorkloadInventory) as tables, build the four pivot
        tables (P0-P3) on 7.PivotTable, and create the two named pivot charts (ChartP0/ChartP1) on
        1.Dashboard that the PowerPoint step copies onto the dashboard slide.

    All COM work is done in a local temp copy (to avoid OneDrive/Protected-View hangs) and the finished
    workbook is copied to the requested destination. The specific Excel instance spawned is tracked by
    PID and terminated in cleanup (only that /automation instance is ever touched).
#>

# --- Excel COM enum constants ---------------------------------------------------------------------
$script:xlRowField    = 1
$script:xlColumnField = 2
$script:xlPageField   = 3
$script:xlCount       = -4112
$script:xlSum         = -4157
$script:xlDatabase    = 1
$script:xlBarClustered = 57
$script:xlLegendTop   = -4160

# In-memory cache of Expert-Analysis worksheet reads (populated by Initialize-WARAExpertAnalysisCache).
$script:WARAExcelReadCache = @{}

function New-WARAExcelApp {
    # Starts a hidden Excel automation instance with all interactive prompts suppressed and returns
    # @{ App = <comobject>; OwnPid = <int> } so the caller can terminate exactly the spawned process.
    $before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $app = New-Object -ComObject Excel.Application
    $ownPid = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.Id -notin $before } | Select-Object -ExpandProperty Id) | Select-Object -First 1
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $app.ScreenUpdating = $false
    try { $app.AutomationSecurity = 1 } catch {}
    try { $app.AskToUpdateLinks = $false } catch {}
    try { $app.AlertBeforeOverwriting = $false } catch {}
    try { $app.EnableEvents = $false } catch {}
    try { $app.Interactive = $false } catch {}
    try { $app.FeatureInstall = 0 } catch {}
    return @{ App = $app; OwnPid = $ownPid }
}

function Remove-WARAExcelApp {
    param($Handle)
    if (-not $Handle) { return }
    $app = $Handle.App
    try { $app.ScreenUpdating = $true } catch {}
    try { $app.Quit() } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {}
    $app = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # Guarantee the spawned instance is gone (only if it is still an /automation EXCEL.EXE).
    if ($Handle.OwnPid) {
        $stillOurs = Get-CimInstance Win32_Process -Filter "ProcessId=$($Handle.OwnPid)" -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -eq 'EXCEL.EXE' -and $_.CommandLine -like '*/automation*' }
        if ($stillOurs) { Stop-Process -Id $Handle.OwnPid -Force -ErrorAction SilentlyContinue }
    }
}

function Read-WARAWorksheetObjects {
    # Reads one worksheet (already-open COM workbook) from a header row into PSCustomObjects.
    param(
        $Workbook,
        [string]$WorksheetName,
        [int]$StartRow,
        [string[]]$AsText = @()
    )
    $ws = $null
    foreach ($w in @($Workbook.Worksheets)) { if ($w.Name -eq $WorksheetName) { $ws = $w; break } }
    if ($null -eq $ws) { throw "Worksheet '$WorksheetName' not found." }

    $used = $ws.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    $lastCol = $used.Column + $used.Columns.Count - 1

    # Header names come from $StartRow. Stop at the first blank header cell.
    $headers = @()
    for ($c = 1; $c -le $lastCol; $c++) {
        $h = $ws.Cells.Item($StartRow, $c).Value2
        if ($null -eq $h -or [string]::IsNullOrEmpty([string]$h)) { break }
        $headers += , @{ Index = $c; Name = [string]$h }
    }

    $rows = New-Object System.Collections.Generic.List[object]
    for ($r = $StartRow + 1; $r -le $lastRow; $r++) {
        $o = [ordered]@{}
        $any = $false
        foreach ($h in $headers) {
            if ($AsText -contains $h.Name) {
                $val = $ws.Cells.Item($r, $h.Index).Text
            }
            else {
                $val = $ws.Cells.Item($r, $h.Index).Value2
            }
            if ($null -ne $val -and -not [string]::IsNullOrEmpty([string]$val)) { $any = $true }
            $o[$h.Name] = $val
        }
        if ($any) { $rows.Add([PSCustomObject]$o) }
    }
    return , $rows.ToArray()
}

function Initialize-WARAExpertAnalysisCache {
    # Opens the Expert-Analysis workbook once (read-only) and caches the sheets the report reads.
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object[]]$Sheets   # each: @{ Name=; StartRow=; AsText=@() }
    )
    $script:WARAExcelReadCache = @{}
    Unblock-File -LiteralPath $Path -ErrorAction SilentlyContinue
    $handle = New-WARAExcelApp
    try {
        $wb = $handle.App.Workbooks.Open($Path, 0, $true)   # ReadOnly
        foreach ($s in $Sheets) {
            $asText = if ($s.ContainsKey('AsText')) { @($s.AsText) } else { @() }
            try {
                $script:WARAExcelReadCache[$s.Name] = Read-WARAWorksheetObjects -Workbook $wb -WorksheetName $s.Name -StartRow $s.StartRow -AsText $asText
            }
            catch {
                # Optional sheets (e.g. Platform Issues / Support) may be absent; cache empty and continue.
                $script:WARAExcelReadCache[$s.Name] = @()
            }
        }
        try { $wb.Close($false) } catch {}
    }
    finally {
        Remove-WARAExcelApp -Handle $handle
    }
}

function Import-WARAExcelViaCom {
    # Drop-in replacement for: Import-Excel -Path <p> -WorksheetName <n> -StartRow <r> [-AsText <c>]
    # Returns cached data when Initialize-WARAExpertAnalysisCache has been called for that sheet.
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$WorksheetName,
        [int]$StartRow = 1,
        [string[]]$AsText = @()
    )
    if ($script:WARAExcelReadCache.ContainsKey($WorksheetName)) {
        return , @($script:WARAExcelReadCache[$WorksheetName])
    }
    Unblock-File -LiteralPath $Path -ErrorAction SilentlyContinue
    $handle = New-WARAExcelApp
    try {
        $wb = $handle.App.Workbooks.Open($Path, 0, $true)
        $data = Read-WARAWorksheetObjects -Workbook $wb -WorksheetName $WorksheetName -StartRow $StartRow -AsText $AsText
        try { $wb.Close($false) } catch {}
        return , @($data)
    }
    finally {
        Remove-WARAExcelApp -Handle $handle
    }
}

function Write-WARAFindingsSheet {
    # Writes header + data rows into an existing worksheet and wraps them in a ListObject table.
    param(
        $Worksheet,
        [int]$StartRow,
        [string[]]$Columns,
        [object[]]$Rows,
        [string]$TableName,
        [string]$TableStyle = 'TableStyleLight19'
    )
    $colCount = $Columns.Count
    $dataCount = @($Rows).Count
    if ($dataCount -lt 1) { $dataCount = 1 }

    # Clear a generous region so no stale template content remains under the table.
    $clearLast = $StartRow + [Math]::Max($dataCount, 5000)
    $Worksheet.Range($Worksheet.Cells.Item($StartRow, 1), $Worksheet.Cells.Item($clearLast, $colCount)).ClearContents() | Out-Null

    # Header row.
    for ($c = 0; $c -lt $colCount; $c++) {
        $Worksheet.Cells.Item($StartRow, $c + 1).Value2 = ($Columns[$c] -replace '_x000a_', "`n")
    }

    # Data block (use .Value, not .Value2, to marshal the 2D object[,] reliably in PowerShell COM).
    if (@($Rows).Count -ge 1) {
        $block = New-Object 'object[,]' @($Rows).Count, $colCount
        for ($r = 0; $r -lt @($Rows).Count; $r++) {
            $row = $Rows[$r]
            for ($c = 0; $c -lt $colCount; $c++) {
                $name = $Columns[$c]
                $val = if ($row -is [System.Collections.IDictionary]) { $row[$name] } else { $row.$name }
                if ($null -ne $val -and $val -isnot [string]) { $val = [string]$val }
                $block[$r, $c] = $val
            }
        }
        $topLeft = $Worksheet.Cells.Item($StartRow + 1, 1)
        $botRight = $Worksheet.Cells.Item($StartRow + @($Rows).Count, $colCount)
        $Worksheet.Range($topLeft, $botRight).Value = $block
    }

    # Table over header + data.
    $lastRow = $StartRow + $dataCount
    $tblRange = $Worksheet.Range($Worksheet.Cells.Item($StartRow, 1), $Worksheet.Cells.Item($lastRow, $colCount))
    foreach ($existing in @($Worksheet.ListObjects)) { if ($existing.Name -eq $TableName) { $existing.Unlist() } }
    $lo = $Worksheet.ListObjects.Add(1, $tblRange, $null, 1)   # xlSrcRange, headers = xlYes
    $lo.Name = $TableName
    try { $lo.TableStyle = $TableStyle } catch {}
    return $lo
}

function Add-WARAPivot {
    # Creates one pivot table from a source Range onto a destination cell, with row/col/page/data config.
    param(
        $Workbook, $DestCell, [string]$Name, $SourceRange,
        [string[]]$Rows = @(), [string[]]$Columns = @(), [string]$PageField,
        [string]$DataField, [int]$DataFunction, [string]$Style
    )
    $cache = $Workbook.PivotCaches().Create($script:xlDatabase, $SourceRange)
    $pt = $cache.CreatePivotTable($DestCell, $Name)
    if ($PageField) { $pt.PivotFields($PageField).Orientation = $script:xlPageField }
    foreach ($r in $Rows) { $pt.PivotFields($r).Orientation = $script:xlRowField }
    foreach ($c in $Columns) { $pt.PivotFields($c).Orientation = $script:xlColumnField }
    $label = ($(if ($DataFunction -eq $script:xlSum) { 'Sum of ' } else { 'Count of ' })) + $DataField
    $null = $pt.AddDataField($pt.PivotFields($DataField), $label, $DataFunction)
    try { $pt.TableStyle2 = $Style } catch {}
    return $pt
}

function Add-WARAPivotChart {
    # Creates a named pivot chart (ChartObject) on a worksheet, bound to a pivot table.
    param($Worksheet, $PivotTable, [string]$Name, [string]$Title, [double]$Left, [double]$Top, [double]$Width, [double]$Height)
    $co = $Worksheet.ChartObjects().Add($Left, $Top, $Width, $Height)
    $co.Name = $Name
    $chart = $co.Chart
    $chart.SetSourceData($PivotTable.TableRange1)
    try { $chart.ChartType = $script:xlBarClustered } catch {}
    try {
        $chart.HasTitle = $true
        $chart.ChartTitle.Text = $Title
    } catch {}
    try { $chart.HasLegend = $true; $chart.Legend.Position = $script:xlLegendTop } catch {}
    # Light grey chart/plot area to match the original dashboard styling (cosmetic).
    $grey = 194 + (194 * 256) + (194 * 65536)
    try { $chart.ChartArea.Format.Fill.ForeColor.RGB = $grey } catch {}
    try { $chart.PlotArea.Format.Fill.ForeColor.RGB = $grey } catch {}
    return $co
}

function Build-WARAAssessmentFindingsCom {
    <#
        Builds the Assessment-Findings workbook from the shipped template using Excel COM:
        data sheets + tables, pivot tables P0-P3, and pivot charts ChartP0/ChartP1. Returns the path.
    #>
    param(
        [Parameter(Mandatory)][string]$TemplatePath,
        [Parameter(Mandatory)][string]$OutputPath,
        [object[]]$ImpactedResources = @(),
        [object[]]$Recommendations = @(),
        [object[]]$WorkloadInventory = @()
    )

    $colsImpacted = @('Impacted?','Resource Type','subscriptionId','resourceGroup','location','name','id','custom1','custom2','custom3','custom4','custom5','Recommendation Title','Impact','Recommendation Control','Potential Benefit','Learn More Link','Long Description','Guid','Category','Source','WAF Pillar','Platform Issue TrackingId','Retirement TrackingId','Support Request Number','Notes','checkName')
    $colsReco     = @('Impact','Description','Potential Benefit','Impacted Resources','Resource Type','Recommendation Control','Long Description','Category','Learn More Link','Guid','Notes')
    $colsInv      = @('id','name','type','tenantId','kind','location','resourceGroup','subscriptionId','managedBy','sku','plan','zones')

    # Build in a local temp copy to avoid OneDrive/Protected-View issues, then copy to $OutputPath.
    $workPath = Join-Path ([System.IO.Path]::GetTempPath()) ('wara-findings-' + [guid]::NewGuid().ToString('N') + '.xlsx')
    Copy-Item -LiteralPath $TemplatePath -Destination $workPath -Force
    Unblock-File -LiteralPath $workPath -ErrorAction SilentlyContinue

    $handle = New-WARAExcelApp
    try {
        $wb = $handle.App.Workbooks.Open($workPath, 0, $false)

        $wsImpacted = $wb.Worksheets.Item('3.ImpactedResources')
        Write-WARAFindingsSheet -Worksheet $wsImpacted -StartRow 12 -Columns $colsImpacted -Rows @($ImpactedResources) -TableName 'impactedresources' | Out-Null

        $wsReco = $wb.Worksheets.Item('2.Recommendations')
        $recoTable = Write-WARAFindingsSheet -Worksheet $wsReco -StartRow 11 -Columns $colsReco -Rows @($Recommendations) -TableName 'recommendationT'

        $wsInv = $wb.Worksheets.Item('6.WorkloadInventory')
        Write-WARAFindingsSheet -Worksheet $wsInv -StartRow 12 -Columns $colsInv -Rows @($WorkloadInventory) -TableName 'WorkloadResources' | Out-Null

        # Pivot tables on 7.PivotTable, sourced from the 2.Recommendations table.
        $wsPivot = $wb.Worksheets.Item('7.PivotTable')
        $src = $recoTable.Range
        $p0 = Add-WARAPivot -Workbook $wb -DestCell $wsPivot.Range('A3') -Name 'P0' -SourceRange $src -Rows @('Resource Type') -Columns @('Impact') -PageField 'Category' -DataField 'Resource Type' -DataFunction $script:xlCount -Style 'PivotStyleMedium9'
        $p1 = Add-WARAPivot -Workbook $wb -DestCell $wsPivot.Range('H3') -Name 'P1' -SourceRange $src -Rows @('Recommendation Control') -Columns @('Impact') -PageField 'Resource Type' -DataField 'Resource Type' -DataFunction $script:xlCount -Style 'PivotStyleMedium9'
        $null = Add-WARAPivot -Workbook $wb -DestCell $wsPivot.Range('O3') -Name 'P2' -SourceRange $src -Rows @('Impact') -DataField 'Impacted Resources' -DataFunction $script:xlSum -Style 'PivotStyleMedium9'
        $null = Add-WARAPivot -Workbook $wb -DestCell $wsPivot.Range('S3') -Name 'P3' -SourceRange $src -Rows @('Impact') -DataField 'Guid' -DataFunction $script:xlCount -Style 'PivotStyleMedium10'

        # Pivot charts on 1.Dashboard (named ChartP0/ChartP1 for the PowerPoint step to copy).
        $wsDash = $wb.Worksheets.Item('1.Dashboard')
        $null = Add-WARAPivotChart -Worksheet $wsDash -PivotTable $p0 -Name 'ChartP0' -Title 'Recommendations by Impact per ResourceType' -Left 10 -Top 10 -Width 450 -Height 525
        $null = Add-WARAPivotChart -Worksheet $wsDash -PivotTable $p1 -Name 'ChartP1' -Title 'Recommendations by Impact per Category' -Left 480 -Top 10 -Width 375 -Height 525

        $handle.App.Calculate()
        $wb.SaveAs($workPath, 51)   # xlOpenXMLWorkbook
        $wb.Close($true)
    }
    finally {
        Remove-WARAExcelApp -Handle $handle
    }

    Copy-Item -LiteralPath $workPath -Destination $OutputPath -Force
    Remove-Item -LiteralPath $workPath -Force -ErrorAction SilentlyContinue
    return $OutputPath
}
