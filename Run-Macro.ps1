Param(
    [string]$excel_file_path,
    [string]$macro_name
)

if ([String]::IsNullOrEmpty($excel_file_path) -or [String]::IsNullOrEmpty($macro_name)) {
    Write-Host 'Usage: .\Run-Macro.ps1 $excel_file_path $macro_name'
    exit 1
}

if (!(Test-Path $excel_file_path)) {
    Write-Host Error: $excel_file_path not found.
    exit 1
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$full_path = Convert-Path $excel_file_path
$book = $excel.Workbooks.Open($full_path)
$excel.run($macro_name)
$book.close()
$book = $null
$excel.quit()
$excel = $null
[GC]::Collect()

exit 0
