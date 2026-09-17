param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$resolvedOutput = (Resolve-Path -LiteralPath $OutputDirectory).Path
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null
try {
    $workbook = $excel.Workbooks.Open($resolvedWorkbook, 0, $true)
    foreach ($sheet in $workbook.Worksheets) {
        if ($sheet.Visible -eq -1) {
            $sheet.ExportAsFixedFormat(0, (Join-Path $resolvedOutput ($sheet.Name + '.pdf')))
            Write-Output $sheet.Name
        }
    }
} finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
