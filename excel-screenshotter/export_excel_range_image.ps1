<#
Example:

Export the used range of a worksheet:
pwsh -File .\export_excel_range_image.ps1 -WorkbookPath "C:\temp\demo.xlsx" -SheetName "Sheet1" -ExportType Range -OutputPath "C:\temp\sheet1.png"

Export a specific range:
pwsh -File .\export_excel_range_image.ps1 -WorkbookPath "C:\temp\demo.xlsx" -SheetName "Sheet1" -ExportType Range -RangeAddress "A1:K30" -OutputPath "C:\temp\range.png"

Export the first Excel table on a worksheet:
pwsh -File .\export_excel_range_image.ps1 -WorkbookPath "C:\temp\dashboard.xlsx" -SheetName "Export" -ExportType Table -OutputPath "C:\temp\table.png"

Export a chart sheet or the first embedded chart on a worksheet:
pwsh -File .\export_excel_range_image.ps1 -WorkbookPath "C:\Users\wu.xiaojun\.openclaw\workspace\excel-master\outputs\china_population_line_chart.xlsx" -SheetName "Population" -ExportType Chart -OutputPath "C:\temp\chart.png"

List objects on a worksheet:
pwsh -File .\export_excel_range_image.ps1 -WorkbookPath "C:\temp\dashboard.xlsx" -SheetName "Export" -ListObjects

JSON output for skill integration:
pwsh -File .\export_excel_range_image.ps1 -WorkbookPath "C:\temp\dashboard.xlsx" -SheetName "Export" -ListObjects -Json
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$SheetName,

    [string]$RangeAddress,

    [ValidateSet('Auto', 'Chart', 'Table', 'Range')]
    [string]$ExportType = 'Auto',

    [string]$TableName,

    [string]$OutputPath,

    [Alias('ListObjects')]
    [switch]$ListMode,

    [switch]$Json,

    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$xlScreen = 1
$xlBitmap = 2

function Release-ComObject {
    param(
        [Parameter(ValueFromPipeline = $true)]
        $ComObject
    )

    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function Get-RangeDescription {
    param(
        $Range
    )

    if ($null -eq $Range) {
        return '<none>'
    }

    return "$($Range.Address($false, $false)) ($($Range.Rows.Count) rows x $($Range.Columns.Count) cols)"
}

function Get-ColumnMetrics {
    param(
        $Range
    )

    $metrics = @()

    if ($null -eq $Range) {
        return $metrics
    }

    for ($i = 1; $i -le $Range.Columns.Count; $i++) {
        $columnRange = $Range.Columns.Item($i)
        try {
            $metrics += "$($columnRange.Address($false, $false)) | Width: $([math]::Round([double]$columnRange.Width, 2))"
        }
        finally {
            Release-ComObject $columnRange
        }
    }

    return $metrics
}

function Get-RowMetrics {
    param(
        $Range
    )

    $metrics = @()

    if ($null -eq $Range) {
        return $metrics
    }

    for ($i = 1; $i -le $Range.Rows.Count; $i++) {
        $rowRange = $Range.Rows.Item($i)
        try {
            $metrics += "$($rowRange.Address($false, $false)) | Height: $([math]::Round([double]$rowRange.Height, 2))"
        }
        finally {
            Release-ComObject $rowRange
        }
    }

    return $metrics
}

function Write-Result {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Result
    )

    if ($Json) {
        $Result | ConvertTo-Json -Depth 8 -Compress
        return
    }

    if ($Result['mode'] -eq 'list') {
        Write-Output "SheetType: $($Result['sheet_type'])"
        Write-Output "SheetName: $($Result['sheet_name'])"
        Write-Output "UsedRange: $($Result['used_range'])"

        if ($Result.ContainsKey('used_range_columns')) {
            Write-Output 'UsedRangeColumns:'
            foreach ($metric in $Result['used_range_columns']) {
                Write-Output "  $metric"
            }
        }

        if ($Result.ContainsKey('used_range_rows')) {
            Write-Output 'UsedRangeRows:'
            foreach ($metric in $Result['used_range_rows']) {
                Write-Output "  $metric"
            }
        }

        Write-Output "Charts: $($Result['charts'].Count)"
        foreach ($chartEntry in $Result['charts']) {
            Write-Output "Chart[$($chartEntry['index'])]: $($chartEntry['name'])"
        }

        Write-Output "Tables: $($Result['tables'].Count)"
        foreach ($tableEntry in $Result['tables']) {
            Write-Output "Table[$($tableEntry['index'])]: $($tableEntry['name']) | Range: $($tableEntry['range'])"
        }

        return
    }

    Write-Output "Image exported to $($Result['output_path'])"
}

$excel = $null
$workbook = $null
$chartSheets = $null
$chartSheet = $null
$worksheet = $null
$worksheets = $null
$listObjects = $null
$listObject = $null
$targetRange = $null
$chartObjects = $null
$chartObject = $null
$chart = $null

try {
    $resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path

    if (-not $ListMode -and [string]::IsNullOrWhiteSpace($OutputPath)) {
        throw 'OutputPath is required unless -ListObjects is used.'
    }

    if (-not $ListMode) {
        $resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
        $outputDirectory = Split-Path -Path $resolvedOutputPath -Parent

        if (-not (Test-Path -LiteralPath $outputDirectory)) {
            [void](New-Item -ItemType Directory -Path $outputDirectory -Force)
        }
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$Visible
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $true

    $workbook = $excel.Workbooks.Open($resolvedWorkbookPath)
    $chartSheets = $workbook.Charts

    $chartSheetExists = $false
    foreach ($sheet in $chartSheets) {
        try {
            if ($sheet.Name -eq $SheetName) {
                $chartSheetExists = $true
                break
            }
        }
        finally {
            Release-ComObject $sheet
        }
    }

    if ($chartSheetExists) {
        if ($ListMode) {
            $chartSheet = $chartSheets.Item($SheetName)
            Write-Result @{
                mode = 'list'
                sheet_type = 'ChartSheet'
                sheet_name = $SheetName
                used_range = '<not applicable>'
                charts = @(@{ index = 1; name = "$SheetName (chart sheet)" })
                tables = @()
            }
            return
        }

        if ($ExportType -ne 'Auto' -and $ExportType -ne 'Chart') {
            throw "Sheet '$SheetName' is a chart sheet. Use -ExportType Chart for this sheet."
        }

        $chartSheet = $chartSheets.Item($SheetName)
        if (-not $chartSheet.Export($resolvedOutputPath, 'PNG')) {
            throw "Excel failed to export chart sheet '$SheetName' to '$resolvedOutputPath'."
        }

        Write-Result @{
            mode = 'export'
            output_path = $resolvedOutputPath
            requested_export_type = $ExportType
            resolved_export_type = 'Chart'
            sheet_type = 'ChartSheet'
            sheet_name = $SheetName
        }
        return
    }

    $worksheets = $workbook.Worksheets

    $worksheetExists = $false
    foreach ($sheet in $worksheets) {
        try {
            if ($sheet.Name -eq $SheetName) {
                $worksheetExists = $true
                break
            }
        }
        finally {
            Release-ComObject $sheet
        }
    }

    if (-not $worksheetExists) {
        throw "Worksheet not found: $SheetName"
    }

    $worksheet = $worksheets.Item($SheetName)
    $worksheet.Activate() | Out-Null
    $chartObjects = $worksheet.ChartObjects()
    $listObjects = $worksheet.ListObjects()

    if ($ListMode) {
        $usedRange = $worksheet.UsedRange
        $chartResults = @()

        for ($i = 1; $i -le $chartObjects.Count; $i++) {
            $chartObjectEntry = $chartObjects.Item($i)
            try {
                $chartResults += @{ index = $i; name = $chartObjectEntry.Name }
            }
            finally {
                Release-ComObject $chartObjectEntry
            }
        }

        $tableResults = @()

        for ($i = 1; $i -le $listObjects.Count; $i++) {
            $tableEntry = $listObjects.Item($i)
            try {
                $tableResults += @{ index = $i; name = $tableEntry.Name; range = (Get-RangeDescription $tableEntry.Range) }
            }
            finally {
                Release-ComObject $tableEntry
            }
        }

        Write-Result @{
            mode = 'list'
            sheet_type = 'Worksheet'
            sheet_name = $SheetName
            used_range = (Get-RangeDescription $usedRange)
            used_range_columns = @(Get-ColumnMetrics $usedRange)
            used_range_rows = @(Get-RowMetrics $usedRange)
            charts = $chartResults
            tables = $tableResults
        }

        Release-ComObject $usedRange
        return
    }

    if (($ExportType -eq 'Auto' -or $ExportType -eq 'Chart') -and [string]::IsNullOrWhiteSpace($RangeAddress) -and $chartObjects.Count -gt 0) {
        $chartObject = $chartObjects.Item(1)
        $chart = $chartObject.Chart

        if (-not $chart.Export($resolvedOutputPath, 'PNG')) {
            throw "Excel failed to export the first embedded chart on worksheet '$SheetName' to '$resolvedOutputPath'."
        }

        Write-Result @{
            mode = 'export'
            output_path = $resolvedOutputPath
            requested_export_type = $ExportType
            resolved_export_type = 'Chart'
            sheet_type = 'Worksheet'
            sheet_name = $SheetName
            chart_name = $chartObject.Name
        }
        return
    }

    if ($ExportType -eq 'Chart') {
        throw "No chart was found on sheet '$SheetName'."
    }

    if ($ExportType -eq 'Table') {
        if (-not [string]::IsNullOrWhiteSpace($TableName)) {
            $tableExists = $false
            foreach ($table in $listObjects) {
                try {
                    if ($table.Name -eq $TableName) {
                        $tableExists = $true
                        break
                    }
                }
                finally {
                    Release-ComObject $table
                }
            }

            if (-not $tableExists) {
                throw "Table not found on worksheet '$SheetName': $TableName"
            }

            $listObject = $listObjects.Item($TableName)
            $targetRange = $listObject.Range
        }
        elseif ($listObjects.Count -gt 0) {
            $listObject = $listObjects.Item(1)
            $targetRange = $listObject.Range
        }
        elseif (-not [string]::IsNullOrWhiteSpace($RangeAddress)) {
            $targetRange = $worksheet.Range($RangeAddress)
        }
        else {
            throw "No Excel table was found on worksheet '$SheetName'. Provide -TableName, or use -ExportType Range with -RangeAddress."
        }
    }

    if ($null -eq $targetRange) {
        if ([string]::IsNullOrWhiteSpace($RangeAddress)) {
            $usedRange = $worksheet.UsedRange
            try {
                $targetRange = $worksheet.Range($usedRange.Address($false, $false))
            }
            finally {
                Release-ComObject $usedRange
            }
        }
        else {
            $targetRange = $worksheet.Range($RangeAddress)
        }
    }

    if ($null -eq $targetRange -or $targetRange.Count -eq 0) {
        throw "No cells were selected for export."
    }

    if (($targetRange.Width -le 0) -or ($targetRange.Height -le 0)) {
        throw "The selected range has no visible size."
    }

    $workbook.Activate() | Out-Null
    $worksheet.Activate() | Out-Null
    $targetRange.Select() | Out-Null
    $excel.CalculateFull()
    $targetRange.CopyPicture($xlScreen, $xlBitmap)

    $chartObject = $chartObjects.Add(0, 0, $targetRange.Width, $targetRange.Height)
    $chart = $chartObject.Chart
    $chart.Paste() | Out-Null

    if (-not $chart.Export($resolvedOutputPath, 'PNG')) {
        throw "Excel failed to export the image to '$resolvedOutputPath'."
    }

    Write-Result @{
        mode = 'export'
        output_path = $resolvedOutputPath
        requested_export_type = $ExportType
        resolved_export_type = $(if ($ExportType -eq 'Table') { 'Table' } else { 'Range' })
        sheet_type = 'Worksheet'
        sheet_name = $SheetName
        range_address = $targetRange.Address($false, $false)
    }
}
finally {
    if ($null -ne $chartObject) {
        try {
            $chartObject.Delete()
        }
        catch {
        }
    }

    if ($null -ne $workbook) {
        $workbook.Close($false)
    }

    if ($null -ne $excel) {
        $excel.Quit()
    }

    $chart | Release-ComObject
    $chartObject | Release-ComObject
    $chartObjects | Release-ComObject
    $listObject | Release-ComObject
    $listObjects | Release-ComObject
    $targetRange | Release-ComObject
    $chartSheet | Release-ComObject
    $chartSheets | Release-ComObject
    $worksheets | Release-ComObject
    $worksheet | Release-ComObject
    $workbook | Release-ComObject
    $excel | Release-ComObject

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}